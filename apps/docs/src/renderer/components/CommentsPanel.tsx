import {
  memo,
  useEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent,
  type ReactNode,
} from 'react'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { CommentInfo } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { absoluteDate, authorInitials, authorInk } from '../editor/comment-card'
import { revWhen } from '../editor/margin-annotations'
import { IconCheck, IconComment } from './icons'

/** anchor text quoted from the document, to orient comments whose range spans paragraphs (one DOM pass for all threads) */
function collectAnchorTexts(): Map<string, string> {
  const out = new Map<string, string>()
  for (const span of document.querySelectorAll<HTMLElement>(
    '.ProseMirror:not(.mp-page) .doc-comment',
  )) {
    const text = span.textContent ?? ''
    for (const id of (span.dataset.commentIds ?? '').split(' ')) {
      if (id) out.set(id, (out.get(id) ?? '') + text)
    }
  }
  return out
}

function jumpTo(id: string) {
  const spans = document.querySelectorAll<HTMLElement>('.ProseMirror:not(.mp-page) .doc-comment')
  const targets = [...spans].filter((s) => (s.dataset.commentIds ?? '').split(' ').includes(id))
  if (targets.length === 0) return
  targets[0].scrollIntoView({ behavior: 'smooth', block: 'center' })
  for (const t of targets) {
    t.classList.remove('doc-comment-flash')
    // restart the CSS animation when jumping to the same comment twice
    void t.offsetWidth
    t.classList.add('doc-comment-flash')
  }
}

/** Word's modern comment header: avatar, author, relative time; thread actions on the right */
function CommentHead({
  c,
  unknownAuthor,
  actions,
}: {
  c: CommentInfo
  unknownAuthor: string
  actions?: ReactNode
}) {
  const name = c.author || unknownAuthor
  return (
    <div className="comment-card-head">
      <span className="comment-avatar" style={{ background: authorInk(name) }}>
        {authorInitials(c)}
      </span>
      <span className="comment-meta">
        <span className="comment-author">{name}</span>
        {c.date && (
          <span className="comment-when" title={absoluteDate(c.date)}>
            {revWhen(c.date)}
          </span>
        )}
      </span>
      {actions}
    </div>
  )
}

/**
 * Word's Comments pane: modern comment cards (avatar, author, relative time,
 * quoted anchor, text, indented replies, inline Reply box), a composer when
 * New Comment was clicked (composing), Resolve and a "…" menu per card.
 */
export const CommentsPanel = memo(function CommentsPanel({
  comments,
  docNode,
  composing,
  onSubmitNew,
  onReply,
  onEdit,
  onResolve,
  onCancelNew,
  onDelete,
  onClose,
  focus,
}: {
  comments: CommentInfo[]
  /** current PM doc, used only as the anchor-scan cache key (doc unchanged ⇒ anchors unchanged) */
  docNode: PmNode
  /** show the new-comment composer (selection captured by the caller) */
  composing: boolean
  onSubmitNew: (text: string) => void
  /** Reply to a comment (anchor shares the parent comment's range) */
  onReply: (parentId: string, text: string) => void
  /** Replace a comment's or reply's text in place; author, date and anchor stay */
  onEdit: (id: string, text: string) => void
  /** Resolve/reopen the whole thread */
  onResolve: (id: string, done: boolean) => void
  onCancelNew: () => void
  onDelete: (id: string) => void
  onClose: () => void
  /** thread to highlight and scroll to (Review > Previous / Next, body anchor click); nonce re-fires
   * for the same id; reply additionally focuses that thread's Reply box */
  focus?: { id: string; nonce: number; reply?: boolean } | null
}) {
  const { t } = useI18n()
  const paneRef = useRef<HTMLElement>(null)
  const [activeId, setActiveId] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const draftRef = useRef<HTMLTextAreaElement>(null)
  const [replyDrafts, setReplyDrafts] = useState<Record<string, string>>({})
  const [editingId, setEditingId] = useState<string | null>(null)
  const [editDraft, setEditDraft] = useState('')
  const [menuId, setMenuId] = useState<string | null>(null)
  // Word: resolved comments are collapsed/hidden by default
  const [showResolved, setShowResolved] = useState(false)

  const anchorTexts = useMemo(collectAnchorTexts, [docNode, comments])

  const topLevel = comments.filter((c) => !c.parentId)
  const openThreads = topLevel.filter((c) => !c.done)
  const resolvedThreads = topLevel.filter((c) => c.done)
  const repliesOf = (id: string) => comments.filter((c) => c.parentId === id)

  const submitReply = (id: string) => {
    const text = (replyDrafts[id] ?? '').trim()
    if (!text) return
    onReply(id, text)
    setReplyDrafts((d) => ({ ...d, [id]: '' }))
  }

  const startEdit = (c: CommentInfo) => {
    // Edit toggles: a second pick cancels instead of resetting the draft
    if (editingId === c.id) {
      setEditingId(null)
      return
    }
    setEditingId(c.id)
    setEditDraft(c.text)
  }

  const submitEdit = () => {
    const text = editDraft.trim()
    if (!text || !editingId) return
    onEdit(editingId, text)
    setEditingId(null)
  }

  /** an open edit belongs to this thread (parent or reply) — commit it before the thread collapses away */
  const flushEditIn = (threadId: string) => {
    if (!editingId) return
    if (editingId !== threadId && !repliesOf(threadId).some((r) => r.id === editingId)) return
    const text = editDraft.trim()
    if (text) onEdit(editingId, text)
    setEditingId(null)
  }

  /** the edited comment is going away — drop the draft with it */
  const cancelEditIn = (ids: string[]) => {
    if (editingId && ids.includes(editingId)) setEditingId(null)
  }

  useEffect(() => {
    if (composing) draftRef.current?.focus()
  }, [composing])

  useEffect(() => {
    if (!menuId) return
    const close = (e: Event) => {
      if (!(e.target instanceof Node) || !paneRef.current?.contains(e.target)) setMenuId(null)
      else if (!(e.target as Element).closest('.comment-menu, .comment-more')) setMenuId(null)
    }
    const onKey = (e: globalThis.KeyboardEvent) => {
      if (e.key === 'Escape') setMenuId(null)
    }
    document.addEventListener('pointerdown', close, true)
    document.addEventListener('keydown', onKey, true)
    return () => {
      document.removeEventListener('pointerdown', close, true)
      document.removeEventListener('keydown', onKey, true)
    }
  }, [menuId])

  useEffect(() => {
    if (!focus) return
    setActiveId(focus.id)
    if (comments.some((c) => c.id === focus.id && c.done)) setShowResolved(true)
    requestAnimationFrame(() => {
      const thread = paneRef.current?.querySelector<HTMLElement>(
        `[data-comment-id="${CSS.escape(focus.id)}"]`,
      )
      thread?.scrollIntoView({ block: 'nearest' })
      if (focus.reply) thread?.querySelector<HTMLInputElement>('.comment-reply-input')?.focus()
    })
    // comments only matter at fire time; re-running on every edit would re-scroll
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [focus])

  const submit = () => {
    const text = draft.trim()
    if (!text) return
    onSubmitNew(text)
    setDraft('')
  }

  const editBox = (
    <div className="comment-reply-compose">
      <textarea
        autoFocus
        value={editDraft}
        onChange={(e) => setEditDraft(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submitEdit()
          if (e.key === 'Escape') setEditingId(null)
        }}
      />
      <div className="comment-compose-actions">
        <button onClick={() => setEditingId(null)}>{t('appCancel')}</button>
        <button className="primary" disabled={!editDraft.trim()} onClick={submitEdit}>
          {t('appSave')}
        </button>
      </div>
    </div>
  )

  const headActions = (c: CommentInfo, root: boolean, replies: CommentInfo[]) => {
    const deleteIds = [c.id, ...replies.map((r) => r.id)]
    const deleteLabel = root
      ? replies.length > 0
        ? t('appDeleteCommentWithReplies')
        : t('appDeleteComment')
      : t('appDeleteReply')
    return (
      <span className="comment-card-actions" onClick={(e) => e.stopPropagation()}>
        {root && (
          <button
            type="button"
            className={`comment-resolve${c.done ? ' on' : ''}`}
            data-tip={c.done ? t('appReopen') : t('appResolve')}
            aria-label={c.done ? t('appReopen') : t('appResolve')}
            onClick={() => {
              // resolving collapses the thread and would unmount an open composer: commit first
              flushEditIn(c.id)
              onResolve(c.id, !c.done)
            }}
          >
            <IconCheck size={13} />
          </button>
        )}
        <button
          type="button"
          className="comment-more"
          aria-label={t('appMoreActions')}
          aria-haspopup="menu"
          aria-expanded={menuId === c.id}
          onClick={() => setMenuId((v) => (v === c.id ? null : c.id))}
        >
          ⋯
        </button>
        {menuId === c.id && (
          <div
            className="comment-menu"
            role="menu"
            ref={(el) => el?.scrollIntoView?.({ block: 'nearest' })}
          >
            <button
              type="button"
              role="menuitem"
              className="comment-card-edit"
              onClick={() => {
                setMenuId(null)
                startEdit(c)
              }}
            >
              {t(root ? 'appEditComment' : 'appEditReply')}
            </button>
            <button
              type="button"
              role="menuitem"
              className="comment-card-del"
              onClick={() => {
                setMenuId(null)
                cancelEditIn(deleteIds)
                onDelete(c.id)
              }}
            >
              {deleteLabel}
            </button>
          </div>
        )}
      </span>
    )
  }

  const onCardKey = (id: string) => (e: KeyboardEvent<HTMLDivElement>) => {
    if (e.target !== e.currentTarget) return
    if (e.key !== 'Enter' && e.key !== ' ') return
    e.preventDefault()
    setActiveId(id)
    jumpTo(id)
  }

  const renderThread = (c: CommentInfo) => {
    const anchor = anchorTexts.get(c.id) ?? ''
    const replies = repliesOf(c.id)
    const editing = editingId === c.id
    const unknown = t('appUnknownAuthor')
    return (
      <div
        key={c.id}
        data-comment-id={c.id}
        className={`comment-thread${c.done ? ' resolved' : ''}${activeId === c.id ? ' active' : ''}`}
      >
        <div
          className={`comment-card${activeId === c.id ? ' active' : ''}`}
          role={editing ? undefined : 'button'}
          tabIndex={editing ? undefined : 0}
          onClick={
            editing
              ? undefined
              : () => {
                  setActiveId(c.id)
                  jumpTo(c.id)
                }
          }
          onKeyDown={editing ? undefined : onCardKey(c.id)}
        >
          <CommentHead c={c} unknownAuthor={unknown} actions={headActions(c, true, replies)} />
          {anchor && <div className="comment-anchor">“{anchor.slice(0, 60)}”</div>}
          {editing ? editBox : <div className="comment-text">{c.text}</div>}
        </div>
        {replies.map((r) => (
          <div key={r.id} className="comment-reply">
            <CommentHead c={r} unknownAuthor={unknown} actions={headActions(r, false, [])} />
            {editingId === r.id ? editBox : <div className="comment-text">{r.text}</div>}
          </div>
        ))}
        {!c.done && (
          <input
            className="comment-reply-input"
            type="text"
            placeholder={t('appReplyPlaceholder')}
            aria-label={t('appReply')}
            value={replyDrafts[c.id] ?? ''}
            onChange={(e) => setReplyDrafts((d) => ({ ...d, [c.id]: e.target.value }))}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
                e.preventDefault()
                submitReply(c.id)
              }
              if (e.key === 'Escape') e.currentTarget.blur()
            }}
          />
        )}
      </div>
    )
  }

  return (
    <aside className="comments-pane" ref={paneRef}>
      <div className="comments-pane-head">
        <span className="comments-pane-title">
          <IconComment size={14} /> {t('appCommentsTitle', { n: comments.length })}
        </span>
        <button
          className="comments-pane-close"
          data-tip={t('appClose')}
          aria-label={t('appClose')}
          onClick={onClose}
        >
          ✕
        </button>
      </div>
      {composing && (
        <div className="comment-compose">
          <textarea
            ref={draftRef}
            placeholder={t('appCommentPlaceholder')}
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit()
              if (e.key === 'Escape') onCancelNew()
            }}
          />
          <div className="comment-compose-actions">
            <button onClick={onCancelNew}>{t('appCancel')}</button>
            <button className="primary" disabled={!draft.trim()} onClick={submit}>
              {t('appCommentBtn')}
            </button>
          </div>
        </div>
      )}
      <div className="comments-pane-list">
        {openThreads.map(renderThread)}
        {resolvedThreads.length > 0 && (
          <div className="comments-resolved-group">
            <button className="comments-resolved-toggle" onClick={() => setShowResolved((v) => !v)}>
              <span className={`comments-resolved-caret ${showResolved ? 'open' : ''}`}>▸</span>
              {t('appResolvedComments', { n: resolvedThreads.length })}
            </button>
            {showResolved && resolvedThreads.map(renderThread)}
          </div>
        )}
        {comments.length === 0 && !composing && (
          <div className="comments-empty">{t('appNoComments')}</div>
        )}
      </div>
    </aside>
  )
})
