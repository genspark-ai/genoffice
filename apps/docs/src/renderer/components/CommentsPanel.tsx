import { memo, useEffect, useMemo, useRef, useState } from 'react'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { CommentInfo } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { IconComment, IconTrash } from './icons'

function formatDate(iso?: string): string {
  if (!iso) return ''
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  return d.toLocaleString(undefined, {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  })
}

/** anchor text quoted from the document, to orient comments whose range spans paragraphs (one DOM pass for all threads) */
function collectAnchorTexts(): Map<string, string> {
  const out = new Map<string, string>()
  for (const span of document.querySelectorAll<HTMLElement>('.ProseMirror .doc-comment')) {
    const text = span.textContent ?? ''
    for (const id of (span.dataset.commentIds ?? '').split(' ')) {
      if (id) out.set(id, (out.get(id) ?? '') + text)
    }
  }
  return out
}

function jumpTo(id: string) {
  const spans = document.querySelectorAll<HTMLElement>('.ProseMirror .doc-comment')
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

/**
 * Word's comments pane: list + jump-to-anchor, a composer when New Comment was
 * clicked (composing), and per-comment delete.
 */
export const CommentsPanel = memo(function CommentsPanel({
  comments,
  docNode,
  composing,
  onSubmitNew,
  onReply,
  onResolve,
  onCancelNew,
  onDelete,
  onClose,
}: {
  comments: CommentInfo[]
  /** current PM doc, used only as the anchor-scan cache key (doc unchanged ⇒ anchors unchanged) */
  docNode: PmNode
  /** show the new-comment composer (selection captured by the caller) */
  composing: boolean
  onSubmitNew: (text: string) => void
  /** Reply to a comment (anchor shares the parent comment's range) */
  onReply: (parentId: string, text: string) => void
  /** Resolve/reopen the whole thread */
  onResolve: (id: string, done: boolean) => void
  onCancelNew: () => void
  onDelete: (id: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [activeId, setActiveId] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const draftRef = useRef<HTMLTextAreaElement>(null)
  const [replyTo, setReplyTo] = useState<string | null>(null)
  const [replyDraft, setReplyDraft] = useState('')
  // Word: resolved comments are collapsed/hidden by default
  const [showResolved, setShowResolved] = useState(false)

  const anchorTexts = useMemo(collectAnchorTexts, [docNode, comments])

  const topLevel = comments.filter((c) => !c.parentId)
  const openThreads = topLevel.filter((c) => !c.done)
  const resolvedThreads = topLevel.filter((c) => c.done)
  const repliesOf = (id: string) => comments.filter((c) => c.parentId === id)
  const submitReply = () => {
    const text = replyDraft.trim()
    if (!text || !replyTo) return
    onReply(replyTo, text)
    setReplyDraft('')
    setReplyTo(null)
  }

  useEffect(() => {
    if (composing) draftRef.current?.focus()
  }, [composing])

  const submit = () => {
    const text = draft.trim()
    if (!text) return
    onSubmitNew(text)
    setDraft('')
  }

  const renderThread = (c: CommentInfo) => {
    const anchor = anchorTexts.get(c.id) ?? ''
    const replies = repliesOf(c.id)
    return (
      <div key={c.id} className={`comment-thread ${c.done ? 'resolved' : ''}`}>
        <button
          className={`comment-card ${activeId === c.id ? 'active' : ''}`}
          onClick={() => {
            setActiveId(c.id)
            jumpTo(c.id)
          }}
        >
          <div className="comment-card-head">
            <span className="comment-avatar">{(c.initials || c.author || '?').slice(0, 2)}</span>
            <span className="comment-author">{c.author || t('appUnknownAuthor')}</span>
            <span className="comment-date">{formatDate(c.date)}</span>
            {c.done && <span className="comment-resolved-badge">{t('appResolved')}</span>}
            <span
              className="comment-card-del"
              title={replies.length > 0 ? t('appDeleteCommentWithReplies') : t('appDeleteComment')}
              role="button"
              aria-label={t('appDeleteComment')}
              tabIndex={0}
              onClick={(e) => {
                e.stopPropagation()
                onDelete(c.id)
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.stopPropagation()
                  onDelete(c.id)
                }
              }}
            >
              <IconTrash size={13} />
            </span>
          </div>
          {anchor && <div className="comment-anchor">“{anchor.slice(0, 60)}”</div>}
          <div className="comment-text">{c.text}</div>
        </button>
        {replies.map((r) => (
          <div key={r.id} className="comment-reply">
            <div className="comment-card-head">
              <span className="comment-avatar">{(r.initials || r.author || '?').slice(0, 2)}</span>
              <span className="comment-author">{r.author || t('appUnknownAuthor')}</span>
              <span className="comment-date">{formatDate(r.date)}</span>
              <span
                className="comment-card-del"
                title={t('appDeleteReply')}
                role="button"
                tabIndex={0}
                onClick={(e) => {
                  e.stopPropagation()
                  onDelete(r.id)
                }}
              >
                <IconTrash size={12} />
              </span>
            </div>
            <div className="comment-text">{r.text}</div>
          </div>
        ))}
        <div className="comment-thread-actions">
          {replyTo === c.id ? (
            <div className="comment-reply-compose">
              <textarea
                autoFocus
                placeholder={t('appReplyPlaceholder')}
                value={replyDraft}
                onChange={(e) => setReplyDraft(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submitReply()
                  if (e.key === 'Escape') setReplyTo(null)
                }}
              />
              <div className="comment-compose-actions">
                <button onClick={() => setReplyTo(null)}>{t('appCancel')}</button>
                <button className="primary" disabled={!replyDraft.trim()} onClick={submitReply}>
                  {t('appReply')}
                </button>
              </div>
            </div>
          ) : (
            <>
              <button onClick={() => setReplyTo(c.id)}>{t('appReply')}</button>
              <button onClick={() => onResolve(c.id, !c.done)}>
                {c.done ? t('appReopen') : t('appResolve')}
              </button>
            </>
          )}
        </div>
      </div>
    )
  }

  return (
    <aside className="comments-pane">
      <div className="comments-pane-head">
        <span className="comments-pane-title">
          <IconComment size={14} /> {t('appCommentsTitle', { n: comments.length })}
        </span>
        <button className="comments-pane-close" title={t('appClose')} onClick={onClose}>
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
