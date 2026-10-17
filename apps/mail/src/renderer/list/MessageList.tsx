import { useEffect, useRef, useState } from 'react'
import type { KeyboardEvent } from 'react'
import type { MailEnvelope } from '@genoffice/mail-protocols'
import { useI18n } from '../i18n/locale'
import {
  LIST_OVERSCAN,
  LIST_ROW_HEIGHT,
  LIST_VIEWPORT_HEIGHT,
  computeListWindow,
} from './list-perf'
import './list.css'

/**
 * Virtualized message list (Lane R owns; H6 perf budgets P-1..P-3).
 * Rows have a fixed height so only the visible window (plus overscan)
 * renders; spacers preserve the scrollbar. Window arithmetic lives in
 * `computeListWindow` (pure O(1): the scroll handler never traverses
 * the list). Arrow keys move the selection; Enter opens via onSelect.
 */

export interface MessageListProps {
  messages: MailEnvelope[]
  selectedUid: string | null
  loading: boolean
  hasMore: boolean
  onSelect(envelope: MailEnvelope): void
  onLoadMore(): void
  rowHeight?: number
  /**
   * H6 test probe for the P-2 render-counter pin: called once per row
   * render with the row uid. Never set in production; off-window rows
   * must never trigger it.
   */
  onRenderRow?(uid: string): void
}

export function MessageList({
  messages,
  selectedUid,
  loading,
  hasMore,
  onSelect,
  onLoadMore,
  rowHeight = LIST_ROW_HEIGHT,
  onRenderRow,
}: MessageListProps) {
  const { t } = useI18n()
  const scrollRef = useRef<HTMLDivElement>(null)
  const [scrollTop, setScrollTop] = useState(0)
  const [viewportHeight, setViewportHeight] = useState(LIST_VIEWPORT_HEIGHT)

  useEffect(() => {
    const node = scrollRef.current
    if (!node) return
    setViewportHeight(node.clientHeight || 480)
    if (typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver((entries) => {
      const height = entries[0]?.contentRect.height
      if (height) setViewportHeight(height)
    })
    observer.observe(node)
    return () => observer.disconnect()
  }, [])

  const window = computeListWindow({
    total: messages.length,
    scrollTop,
    viewportHeight,
    rowHeight,
    overscan: LIST_OVERSCAN,
  })
  const windowed = messages.slice(window.startIndex, window.endIndex)
  const topSpacer = window.topSpacer
  const bottomSpacer = window.bottomSpacer

  const moveSelection = (delta: number): void => {
    if (messages.length === 0) return
    const current = messages.findIndex((item) => item.ref.uid === selectedUid)
    const next = current < 0 ? (delta > 0 ? 0 : messages.length - 1) : current + delta
    const clamped = Math.min(Math.max(next, 0), messages.length - 1)
    const target = messages[clamped]
    if (target) onSelect(target)
  }

  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>): void => {
    if (event.key === 'ArrowDown') {
      event.preventDefault()
      moveSelection(1)
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      moveSelection(-1)
    }
  }

  return (
    <div className="mail-list-wrap">
      <div
        ref={scrollRef}
        className="mail-list"
        role="listbox"
        aria-label={t('mail.accounts.title')}
        aria-busy={loading}
        tabIndex={0}
        onKeyDown={onKeyDown}
        onScroll={(event) => setScrollTop(event.currentTarget.scrollTop)}
      >
        {loading && messages.length === 0 && (
          <div className="mail-list-status">{t('mail.list.loading')}</div>
        )}
        {!loading && messages.length === 0 && (
          <div className="mail-list-status">{t('mail.list.empty')}</div>
        )}
        {topSpacer > 0 && <div style={{ height: topSpacer }} aria-hidden="true" />}
        {windowed.map((envelope) => {
          const active = envelope.ref.uid === selectedUid
          if (onRenderRow) onRenderRow(envelope.ref.uid)
          return (
            <button
              key={`${envelope.ref.folderId}/${envelope.ref.uid}`}
              type="button"
              role="option"
              aria-selected={active}
              className={active ? 'mail-row is-active' : 'mail-row'}
              style={{ height: rowHeight }}
              onClick={() => onSelect(envelope)}
              onDoubleClick={() => onSelect(envelope)}
            >
              <span className="mail-row-main">
                <span className="mail-row-top">
                  {!envelope.flags.seen && (
                    <span className="mail-row-unread" aria-label={t('mail.folders.unreadCount')} />
                  )}
                  <span className="mail-row-from">{envelope.from.join(', ')}</span>
                  <span className="mail-row-date">{envelope.date}</span>
                </span>
                <span className="mail-row-subject">
                  {envelope.flags.flagged && (
                    <span className="mail-row-flag" aria-hidden="true">
                      *
                    </span>
                  )}
                  {envelope.subject}
                </span>
                <span className="mail-row-snippet">{envelope.snippet}</span>
              </span>
              {envelope.hasAttach && (
                <span
                  className="mail-row-attach"
                  aria-label={t('mail.reader.attachments')}
                  aria-hidden="false"
                >
                  @
                </span>
              )}
            </button>
          )
        })}
        {bottomSpacer > 0 && <div style={{ height: bottomSpacer }} aria-hidden="true" />}
      </div>
      <div className="mail-list-footer">
        <span className="mail-list-count" aria-live="polite">
          {messages.length} {t('mail.list.virtualizedCount')}
        </span>
        {hasMore && (
          <button type="button" className="mail-list-more" onClick={onLoadMore}>
            {t('mail.list.moreOnServer')}
          </button>
        )}
      </div>
    </div>
  )
}
