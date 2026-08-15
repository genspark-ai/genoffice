/**
 * Comments pane (right side) — current page's comment list + create/delete.
 * The source of truth lives in the main process (the pptx comments part); this only displays and
 * sends back intents. The list is owned by App (refreshed after page switch/add/delete), keeping
 * the ribbon badge in sync.
 */
import React, { useEffect, useRef, useState } from 'react'
import type { SlideComment } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { IconComment, IconSidebarCollapse } from './icons'

interface Props {
  slideIndex: number
  comments: SlideComment[]
  /** Focus the input when it changes (the "new comment" entry increments it per click) */
  focusNonce?: number
  onAdd: (text: string) => void
  onDelete: (c: SlideComment) => void
  onCollapse: () => void
}

/** dt (ISO) → local short time; returned as-is on parse failure. */
function fmtDt(dt: string, locale: string): string {
  const d = new Date(dt)
  if (Number.isNaN(d.getTime())) return dt
  return d.toLocaleString(locale, {
    month: 'numeric',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}

export function CommentsPane({
  slideIndex,
  comments,
  focusNonce,
  onAdd,
  onDelete,
  onCollapse,
}: Props) {
  const { t, dateLocale } = useI18n()
  const [draft, setDraft] = useState('')
  const inputRef = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    if (focusNonce) inputRef.current?.focus()
  }, [focusNonce])

  const submit = () => {
    const text = draft.trim()
    if (!text) return
    onAdd(text)
    setDraft('')
  }

  return (
    <aside className="comments-pane">
      <div className="ai-panel-header">
        <span className="ai-panel-title">{t('paneCommentsTitle', { n: slideIndex + 1 })}</span>
        <div className="ai-panel-header-actions">
          <button
            className="ai-header-btn"
            onClick={onCollapse}
            data-tip={t('paneCommentsCollapse')}
            aria-label={t('paneCommentsCollapse')}
          >
            <IconSidebarCollapse size={15} />
          </button>
        </div>
      </div>

      <div className="comments-list">
        {comments.length === 0 && (
          <div className="comments-empty">
            <IconComment size={22} />
            <p>{t('paneCommentsEmpty')}</p>
            <p className="comments-empty-sub">{t('paneCommentsEmptySub')}</p>
          </div>
        )}
        {comments.map((c) => (
          <div key={`${c.authorId}-${c.idx}`} className="comment-card">
            <div className="comment-head">
              <span className="comment-avatar">{(c.initials || c.author).slice(0, 2)}</span>
              <span className="comment-meta">
                <span className="comment-author">{c.author}</span>
                {c.dt && <span className="comment-time">{fmtDt(c.dt, dateLocale)}</span>}
              </span>
              <button
                className="comment-del"
                data-tip={t('paneCommentsDelete')}
                aria-label={t('paneCommentsDelete')}
                onClick={() => onDelete(c)}
              >
                ✕
              </button>
            </div>
            <div className="comment-text">{c.text}</div>
          </div>
        ))}
      </div>

      <div className="comment-new">
        <textarea
          ref={inputRef}
          rows={3}
          value={draft}
          placeholder={t('paneCommentsPlaceholder')}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
              e.preventDefault()
              submit()
            }
          }}
        />
        <button className="btn-primary" disabled={!draft.trim()} onClick={submit}>
          {t('paneCommentsPost')}
        </button>
      </div>
    </aside>
  )
}
