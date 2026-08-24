import React, { useEffect, useState } from 'react'

/**
 * Ephemeral thinking/reasoning block: shows live reasoning text while the
 * model thinks, then collapses to "Thought for Ns" when the answer streams.
 * Not persisted — history stays small, same as the interrupted/truncated cards.
 */
export function AiThinkingBlock({
  text,
  done,
  labelThinking,
  labelThoughtFor,
}: {
  readonly text: string
  /** true once the run's final text starts streaming or the turn ends */
  readonly done?: boolean
  readonly labelThinking: string
  readonly labelThoughtFor: (n: number) => string
}): React.JSX.Element | null {
  const [open, setOpen] = useState(true)
  const [elapsed, setElapsed] = useState(0)

  // Track elapsed while thinking is live; freeze when done
  useEffect(() => {
    if (done) return
    if (!text) return
    const start = Date.now()
    const id = window.setInterval(() => setElapsed(Math.round((Date.now() - start) / 1000)), 1000)
    return () => window.clearInterval(id)
  }, [text, done])

  // Auto-collapse when the answer starts arriving (done) — keep manual open as override
  useEffect(() => {
    if (done) setOpen(false)
    else if (text && !done) setOpen(true)
  }, [done, text])

  if (!text) return null

  // While live, show elapsed in the header after 3s (mirrors AiTypingIndicator)
  const header = done
    ? labelThoughtFor(elapsed || 0)
    : elapsed >= 3
      ? `${labelThinking} · ${elapsed}s`
      : labelThinking

  return (
    <div className={`ai-thinking${open ? ' ai-thinking-open' : ''}`}>
      <button
        type="button"
        className="ai-thinking-header"
        onClick={() => setOpen((v) => !v)}
        aria-expanded={open}
      >
        <span className="ai-thinking-chevron" aria-hidden>
          {open ? '▾' : '▸'}
        </span>
        <span className="ai-thinking-title">{header}</span>
      </button>
      {open && <div className="ai-thinking-body">{text}</div>}
    </div>
  )
}
