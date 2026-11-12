import React, { useEffect, useRef, useState } from 'react'
import { IconBroom, IconChevronDown, IconPause, IconPencil, IconPlay, IconTrash } from './icons'
import type { AiQueueStripLabels } from './strings-ai-queue'

/** how long a press has to hold before it turns into a drag */
const DRAG_HOLD_MS = 220
/** movement below this is a press, not a drag */
const DRAG_SLOP_PX = 4

/** the list as it would look with `from` moved to `to` — the drag preview */
export function moveWithin<T>(items: readonly T[], from: number, to: number): T[] {
  if (from < 0 || from >= items.length) return [...items]
  const target = Math.max(0, Math.min(items.length - 1, to))
  if (target === from) return [...items]
  const next = [...items]
  next.splice(target, 0, ...next.splice(from, 1))
  return next
}

/**
 * The queued-messages strip above the AI composer (MiniMax-style): a collapsed
 * "N queued" row that expands to one editable row per message, plus a clear-all
 * button. Rendered inside the composer's input box; the app's CSS themes it
 * like the rest of the `.ai-input-box` family.
 *
 * Pausing holds the queue: the run in flight still finishes, but nothing after
 * it starts until the gate is opened again, so a message can be edited without
 * the next one slipping out. Rows reorder by press-and-drag (or Alt+arrows),
 * which only changes the order — nothing starts because of it.
 */
export function AiQueueStrip({
  items,
  labels,
  paused = false,
  onTogglePause,
  onUpdate,
  onMove,
  onRemove,
  onClear,
}: {
  readonly items: ReadonlyArray<{ readonly id: string; readonly text: string }>
  readonly labels: AiQueueStripLabels
  /** queue gate closed: nothing after the current run starts on its own */
  readonly paused?: boolean
  readonly onTogglePause?: () => void
  readonly onUpdate: (id: string, text: string) => void
  /** put a queued message at another position in the queue */
  readonly onMove?: (id: string, toIndex: number) => void
  readonly onRemove: (id: string) => void
  readonly onClear: () => void
}): React.JSX.Element | null {
  const [open, setOpen] = useState(false)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const editRef = useRef<HTMLInputElement | null>(null)
  const listRef = useRef<HTMLUListElement | null>(null)
  const [drag, setDrag] = useState<{ id: string; overIndex: number } | null>(null)
  // the drag lives in a ref too: the pointer handlers run outside React's render
  const dragRef = useRef<{ id: string; overIndex: number } | null>(null)
  const endPointerRef = useRef<(() => void) | null>(null)

  useEffect(() => {
    if (editingId) editRef.current?.focus()
  }, [editingId])

  useEffect(() => () => endPointerRef.current?.(), [])

  if (items.length === 0) return null

  const setDragBoth = (next: { id: string; overIndex: number } | null) => {
    dragRef.current = next
    setDrag(next)
  }

  /** which slot a pointer Y falls in, by row midpoints */
  const indexAt = (clientY: number) => {
    const rows = Array.from(listRef.current?.querySelectorAll('li.ai-queue-item') ?? [])
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i]!.getBoundingClientRect()
      if (clientY < r.top + r.height / 2) return i
    }
    return rows.length - 1
  }

  /** a press on a row: hold it to pick the row up, then move it up or down */
  const beginPress = (id: string, clientY: number) => {
    if (!onMove) return
    const timer = window.setTimeout(
      () => setDragBoth({ id, overIndex: indexAt(clientY) }),
      DRAG_HOLD_MS,
    )
    const move = (e: PointerEvent) => {
      if (dragRef.current) setDragBoth({ id, overIndex: indexAt(e.clientY) })
      // moving before the hold completes is a scroll or a stray press, not a drag
      else if (Math.abs(e.clientY - clientY) > DRAG_SLOP_PX) end()
    }
    const end = () => {
      const held = dragRef.current
      window.clearTimeout(timer)
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', end)
      window.removeEventListener('pointercancel', end)
      endPointerRef.current = null
      if (held) {
        setDragBoth(null)
        onMove(held.id, held.overIndex)
      }
    }
    endPointerRef.current = end
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', end)
    window.addEventListener('pointercancel', end)
  }

  const commitEdit = () => {
    if (!editingId) return
    const text = draft.trim()
    if (text) onUpdate(editingId, text)
    setEditingId(null)
  }

  const from = drag ? items.findIndex((i) => i.id === drag.id) : -1
  const shown = drag && from >= 0 ? moveWithin(items, from, drag.overIndex) : items

  const onRowKeyDown = (e: React.KeyboardEvent<HTMLLIElement>, id: string, index: number) => {
    if (!e.altKey) return
    if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      e.preventDefault()
      onMove?.(id, index + (e.key === 'ArrowUp' ? -1 : 1))
    }
  }

  return (
    <div className="ai-queue-strip" data-paused={paused ? 'true' : undefined}>
      <div className="ai-queue-head">
        <button
          type="button"
          className="ai-queue-toggle"
          onClick={() => setOpen((v) => !v)}
          aria-expanded={open}
        >
          <IconChevronDown size={12} />
          {labels.queuedCount(items.length)}
        </button>
        {paused && <span className="ai-queue-hint">{labels.pausedHint}</span>}
        {onTogglePause && (
          <button
            type="button"
            className="ai-queue-pause"
            onClick={onTogglePause}
            aria-pressed={paused}
            data-tip={paused ? labels.resumeTitle : labels.pauseTitle}
            aria-label={paused ? labels.resumeTitle : labels.pauseTitle}
          >
            {paused ? <IconPlay size={13} /> : <IconPause size={13} />}
          </button>
        )}
        <button
          type="button"
          className="ai-queue-clear"
          onClick={onClear}
          data-tip={labels.clearTitle}
          aria-label={labels.clearTitle}
        >
          <IconBroom size={13} />
        </button>
      </div>
      {open && (
        <ul className={`ai-queue-items${drag ? ' ai-queue-dragging' : ''}`} ref={listRef}>
          {shown.map((item) => {
            const index = shown.indexOf(item)
            return (
              <li
                key={item.id}
                className="ai-queue-item"
                tabIndex={0}
                data-tip={onMove ? labels.moveTitle : undefined}
                data-dragging={drag?.id === item.id ? 'true' : undefined}
                onKeyDown={(e) => onRowKeyDown(e, item.id, index)}
                onPointerDown={(e) => {
                  // buttons inside the row keep their own click meaning
                  if ((e.target as HTMLElement).closest('button')) return
                  beginPress(item.id, e.clientY)
                }}
              >
                {editingId === item.id ? (
                  <input
                    ref={editRef}
                    className="ai-queue-edit"
                    value={draft}
                    dir="auto"
                    onChange={(e) => setDraft(e.target.value)}
                    onBlur={commitEdit}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
                        e.preventDefault()
                        commitEdit()
                      } else if (e.key === 'Escape') {
                        e.preventDefault()
                        setEditingId(null)
                      }
                    }}
                  />
                ) : (
                  <span className="ai-queue-text" dir="auto">
                    {item.text}
                  </span>
                )}
                {editingId === item.id ? null : (
                  <>
                    <button
                      type="button"
                      className="ai-queue-action"
                      onClick={() => {
                        setDraft(item.text)
                        setEditingId(item.id)
                      }}
                      data-tip={labels.editTitle}
                      aria-label={labels.editTitle}
                    >
                      <IconPencil size={13} />
                    </button>
                    <button
                      type="button"
                      className="ai-queue-action"
                      onClick={() => onRemove(item.id)}
                      data-tip={labels.removeTitle}
                      aria-label={labels.removeTitle}
                    >
                      <IconTrash size={13} />
                    </button>
                  </>
                )}
              </li>
            )
          })}
        </ul>
      )}
    </div>
  )
}
