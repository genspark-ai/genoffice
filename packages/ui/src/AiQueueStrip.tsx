import React, { useEffect, useRef, useState } from 'react'
import { IconBroom, IconChevronDown, IconPause, IconPencil, IconPlay, IconTrash } from './icons'
import type { AiQueueStripLabels } from './strings-ai-queue'

/**
 * The queued-messages strip above the AI composer (MiniMax-style): a collapsed
 * "N queued" row that expands to one editable row per message, plus a clear-all
 * button. Rendered inside the composer's input box; the app's CSS themes it
 * like the rest of the `.ai-input-box` family.
 *
 * Pausing holds the queue: the run in flight still finishes, but nothing after
 * it starts until the gate is opened again, so a message can be edited without
 * the next one slipping out.
 */
export function AiQueueStrip({
  items,
  labels,
  paused = false,
  onTogglePause,
  onUpdate,
  onRemove,
  onClear,
}: {
  readonly items: ReadonlyArray<{ readonly id: string; readonly text: string }>
  readonly labels: AiQueueStripLabels
  /** queue gate closed: nothing after the current run starts on its own */
  readonly paused?: boolean
  readonly onTogglePause?: () => void
  readonly onUpdate: (id: string, text: string) => void
  readonly onRemove: (id: string) => void
  readonly onClear: () => void
}): React.JSX.Element | null {
  const [open, setOpen] = useState(false)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const editRef = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    if (editingId) editRef.current?.focus()
  }, [editingId])

  if (items.length === 0) return null

  const commitEdit = () => {
    if (!editingId) return
    const text = draft.trim()
    if (text) onUpdate(editingId, text)
    setEditingId(null)
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
        <ul className="ai-queue-items">
          {items.map((item) => (
            <li key={item.id} className="ai-queue-item">
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
          ))}
        </ul>
      )}
    </div>
  )
}
