import { useMemo, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import type { Lang } from '@genoffice/i18n'
import type { TFunc } from './i18n/locale'
import { NOTE_FALLBACK_COLOR, cssRgb } from './DrawLayer'
import { flattenThread } from './note-threads'
import type { NoteThreadItem } from './note-threads'
import { CARD_PIN_ALIGN, layoutMarginCards } from './note-margin-layout'

/** Pencil glyph for rewriting an already-confirmed comment */
const IconPencil = (): ReactElement => (
  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path
      d="M4 20h4L19.5 8.5a2.1 2.1 0 0 0-3-3L5 17v3ZM14.5 7.5l2 2"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
)

/** Small trash-can glyph so the per-comment delete cannot be mistaken for a close button */
const IconTrash = (): ReactElement => (
  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path
      d="M4 7h16M10 11v6M14 11v6M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13M9 7V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v3"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
)

/** A comment thread shown as a card in the margin, tied to a pin on a page of this row */
export interface NoteMarginThread {
  origIdx: number
  root: NoteThreadItem
  /** Pin center relative to the margin column top-left (x negative: pins sit on the pages) */
  pinX: number
  pinY: number
}

/** The comment being typed for a freshly placed pin (WPS-style inline creation) */
export interface NoteMarginDraft {
  /** Identity of the placement; changing it remounts the draft card (fresh text/timestamp) */
  placementKey: string
  origIdx: number
  pinX: number
  pinY: number
  color: [number, number, number]
}

const DRAFT_KEY = '\u0000draft'

function threadColor(root: NoteThreadItem): [number, number, number] {
  return root.color ?? NOTE_FALLBACK_COLOR
}

/** Leader line from the pin on the page to the card's attachment row (WPS-style) */
function LeadLine({
  pinX,
  pinY,
  cardTop,
  color,
}: {
  pinX: number
  pinY: number
  cardTop: number
  color: [number, number, number]
}): ReactElement {
  const rgb = cssRgb(color)
  const y2 = cardTop + CARD_PIN_ALIGN
  return (
    <svg className="pdf-note-lead" aria-hidden="true">
      <circle cx={pinX} cy={pinY} r={2.5} fill={rgb} />
      <polyline
        points={`${pinX},${pinY} -10,${y2} 0,${y2}`}
        fill="none"
        stroke={rgb}
        strokeWidth={1.5}
      />
    </svg>
  )
}

function NoteCard({
  thread,
  active,
  top,
  readOnly,
  timeFmt,
  t,
  setEl,
  onActivate,
  onReply,
  editingKey,
  editingText,
  onEditStart,
  onEditChange,
  onEditCancel,
  onEditSubmit,
  onDeleteItem,
  onClose,
}: {
  thread: NoteMarginThread
  active: boolean
  top: number
  readOnly: boolean
  timeFmt: Intl.DateTimeFormat
  t: TFunc
  setEl: (el: HTMLDivElement | null) => void
  onActivate: () => void
  onReply: (text: string) => void
  /** Key of the comment whose body is an edit box; state lives in App so a save can
      fold the in-progress text in before the post-save reload unmounts this card */
  editingKey: string | null
  editingText: string
  onEditStart: (item: NoteThreadItem) => void
  onEditChange: (text: string) => void
  onEditCancel: () => void
  onEditSubmit: (item: NoteThreadItem) => void
  onDeleteItem: (item: NoteThreadItem) => void
  onClose: () => void
}): ReactElement {
  const [text, setText] = useState('')
  const items = flattenThread(thread.root)

  const submit = () => {
    const trimmed = text.trim()
    if (!trimmed) return
    setText('')
    onReply(trimmed)
  }

  return (
    <div
      ref={setEl}
      className={`pdf-note-card${active ? ' pdf-note-card-active' : ''}`}
      style={{ top, borderTopColor: cssRgb(threadColor(thread.root)) }}
      onClick={active ? undefined : onActivate}
    >
      <div className="pdf-note-card-head">
        <span className="pdf-note-card-title">{t('noteThread')}</span>
        {active && (
          <button
            type="button"
            className="pdf-note-card-close"
            data-tip={t('noteClose')}
            aria-label={t('noteClose')}
            onClick={(e) => {
              e.stopPropagation()
              onClose()
            }}
          >
            ×
          </button>
        )}
      </div>
      <div className="pdf-note-thread">
        {items.map(({ item, depth }) => (
          <div
            key={item.key}
            className={`pdf-note-comment${depth > 0 ? ' pdf-note-comment-reply' : ''}`}
          >
            <div className="pdf-note-comment-head">
              <span className="pdf-note-author">{item.author || 'GenOffice'}</span>
              <span className="pdf-note-time">
                {item.timeMs !== null ? timeFmt.format(item.timeMs) : ''}
              </span>
              {active && !readOnly && editingKey !== item.key && (
                <span className="pdf-note-comment-actions">
                  <button
                    type="button"
                    className="pdf-note-comment-edit"
                    data-tip={t('noteEdit')}
                    aria-label={t('noteEdit')}
                    onClick={(e) => {
                      e.stopPropagation()
                      onEditStart(item)
                    }}
                  >
                    <IconPencil />
                  </button>
                  <button
                    type="button"
                    className="pdf-note-comment-del"
                    data-tip={t('deleteAnnotation')}
                    aria-label={t('deleteAnnotation')}
                    onClick={(e) => {
                      e.stopPropagation()
                      onDeleteItem(item)
                    }}
                  >
                    <IconTrash />
                  </button>
                </span>
              )}
            </div>
            {active && !readOnly && editingKey === item.key ? (
              <div className="pdf-note-edit-box">
                <textarea
                  rows={3}
                  value={editingText}
                  autoFocus
                  onChange={(e) => onEditChange(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
                      if (editingText.trim()) onEditSubmit(item)
                    } else if (e.key === 'Escape') {
                      e.stopPropagation()
                      onEditCancel()
                    }
                  }}
                />
                <div className="pdf-note-draft-actions">
                  <button type="button" className="pdf-modal-btn" onClick={onEditCancel}>
                    {t('cancel')}
                  </button>
                  <button
                    type="button"
                    className="pdf-modal-btn primary"
                    disabled={!editingText.trim()}
                    onClick={() => onEditSubmit(item)}
                  >
                    {t('ok')}
                  </button>
                </div>
              </div>
            ) : (
              <div className="pdf-note-comment-body">{item.contents}</div>
            )}
          </div>
        ))}
      </div>
      {active && !readOnly && (
        <div className="pdf-note-reply-box">
          <textarea
            rows={2}
            value={text}
            placeholder={t('noteReplyPlaceholder')}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit()
              else if (e.key === 'Escape') {
                e.stopPropagation()
                onClose()
              }
            }}
          />
          <button
            type="button"
            className="pdf-modal-btn primary pdf-note-reply-send"
            disabled={!text.trim()}
            onClick={submit}
          >
            {t('noteReply')}
          </button>
        </div>
      )}
    </div>
  )
}

/** Empty card with a textarea for the note just placed on the page */
function NoteDraftCard({
  author,
  color,
  top,
  timeFmt,
  t,
  setEl,
  onConfirm,
  onCancel,
}: {
  author: string
  color: [number, number, number]
  top: number
  timeFmt: Intl.DateTimeFormat
  t: TFunc
  setEl: (el: HTMLDivElement | null) => void
  onConfirm: (text: string) => void
  onCancel: () => void
}): ReactElement {
  const [text, setText] = useState('')
  const [createdMs] = useState(() => Date.now())

  const submit = () => {
    const trimmed = text.trim()
    if (trimmed) onConfirm(trimmed)
    else onCancel()
  }

  return (
    <div
      ref={setEl}
      className="pdf-note-card pdf-note-card-active"
      style={{ top, borderTopColor: cssRgb(color) }}
    >
      <div className="pdf-note-card-head">
        <span className="pdf-note-author">{author || 'GenOffice'}</span>
        <span className="pdf-note-time">{timeFmt.format(createdMs)}</span>
      </div>
      <div className="pdf-note-draft-box">
        <textarea
          rows={3}
          value={text}
          placeholder={t('notePlaceholder')}
          autoFocus
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit()
            else if (e.key === 'Escape') {
              e.stopPropagation()
              onCancel()
            }
          }}
        />
        <div className="pdf-note-draft-actions">
          <button type="button" className="pdf-modal-btn" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button
            type="button"
            className="pdf-modal-btn primary"
            disabled={!text.trim()}
            onClick={submit}
          >
            {t('ok')}
          </button>
        </div>
      </div>
    </div>
  )
}

/**
 * WPS-style comments margin beside one page row: every thread of the row shows as a
 * card level with its pin, the active (or draft) card is pinned in place with a leader
 * line to its pin, and the rest yield vertically so cards never overlap. Clicking a
 * card activates its thread (which also highlights the pin on the page).
 */
export function NoteMarginColumn({
  width,
  threads,
  draft,
  activeKey,
  author,
  lang,
  readOnly,
  t,
  onActivate,
  onReply,
  editingKey,
  editingText,
  onEditStart,
  onEditChange,
  onEditCancel,
  onEditSubmit,
  onDeleteItem,
  onClose,
  onDraftConfirm,
  onDraftCancel,
}: {
  width: number
  threads: NoteMarginThread[]
  draft: NoteMarginDraft | null
  /** Root key of the active thread in this row; null when none */
  activeKey: string | null
  /** Default author shown on the draft card */
  author: string
  lang: Lang
  readOnly: boolean
  t: TFunc
  onActivate: (thread: NoteMarginThread) => void
  onReply: (thread: NoteMarginThread, text: string) => void
  /** Comment currently being rewritten (App-held so saves can flush it) */
  editingKey: string | null
  editingText: string
  onEditStart: (thread: NoteMarginThread, item: NoteThreadItem) => void
  onEditChange: (text: string) => void
  onEditCancel: () => void
  onEditSubmit: (thread: NoteMarginThread, item: NoteThreadItem) => void
  onDeleteItem: (thread: NoteMarginThread, item: NoteThreadItem) => void
  onClose: () => void
  onDraftConfirm: (text: string) => void
  onDraftCancel: () => void
}): ReactElement {
  const [heights, setHeights] = useState<Map<string, number>>(new Map())
  const timeFmt = useMemo(
    () => new Intl.DateTimeFormat(lang, { dateStyle: 'medium', timeStyle: 'short' }),
    [lang],
  )

  // Card heights via per-card ResizeObserver, so the layout follows content changes
  // (replies added, thread expanded on activation). Ref callbacks are memoized per
  // key: a fresh closure each render would remount the observer every render.
  const observersRef = useRef(new Map<string, ResizeObserver>())
  const refCbsRef = useRef(new Map<string, (el: HTMLDivElement | null) => void>())
  const setEl = (key: string): ((el: HTMLDivElement | null) => void) => {
    let cb = refCbsRef.current.get(key)
    if (!cb) {
      cb = (el) => {
        observersRef.current.get(key)?.disconnect()
        observersRef.current.delete(key)
        if (el) {
          const ro = new ResizeObserver(() => {
            const h = el.offsetHeight
            setHeights((prev) => (prev.get(key) === h ? prev : new Map(prev).set(key, h)))
          })
          ro.observe(el)
          observersRef.current.set(key, ro)
        } else {
          setHeights((prev) => {
            if (!prev.has(key)) return prev
            const next = new Map(prev)
            next.delete(key)
            return next
          })
        }
      }
      refCbsRef.current.set(key, cb)
    }
    return cb
  }

  const entries = [
    ...threads.map((th) => ({ key: th.root.key, pinY: th.pinY })),
    ...(draft ? [{ key: DRAFT_KEY, pinY: draft.pinY }] : []),
  ]
  const anchorKey = draft ? DRAFT_KEY : activeKey
  const tops = layoutMarginCards(entries, (key) => heights.get(key) ?? 120, anchorKey)

  const activeThread = activeKey !== null ? threads.find((th) => th.root.key === activeKey) : null

  return (
    <div className="pdf-note-margin" style={{ width }}>
      {draft ? (
        <LeadLine
          pinX={draft.pinX}
          pinY={draft.pinY}
          cardTop={tops.get(DRAFT_KEY) ?? 0}
          color={draft.color}
        />
      ) : (
        activeThread && (
          <LeadLine
            pinX={activeThread.pinX}
            pinY={activeThread.pinY}
            cardTop={tops.get(activeThread.root.key) ?? 0}
            color={threadColor(activeThread.root)}
          />
        )
      )}
      {threads.map((th) => (
        <NoteCard
          key={th.root.key}
          thread={th}
          active={th.root.key === activeKey}
          top={tops.get(th.root.key) ?? 0}
          readOnly={readOnly}
          timeFmt={timeFmt}
          t={t}
          setEl={setEl(th.root.key)}
          onActivate={() => onActivate(th)}
          onReply={(text) => onReply(th, text)}
          editingKey={editingKey}
          editingText={editingText}
          onEditStart={(item) => onEditStart(th, item)}
          onEditChange={onEditChange}
          onEditCancel={onEditCancel}
          onEditSubmit={(item) => onEditSubmit(th, item)}
          onDeleteItem={(item) => onDeleteItem(th, item)}
          onClose={onClose}
        />
      ))}
      {draft && (
        <NoteDraftCard
          key={draft.placementKey}
          author={author}
          color={draft.color}
          top={tops.get(DRAFT_KEY) ?? 0}
          timeFmt={timeFmt}
          t={t}
          setEl={setEl(DRAFT_KEY)}
          onConfirm={onDraftConfirm}
          onCancel={onDraftCancel}
        />
      )}
    </div>
  )
}
