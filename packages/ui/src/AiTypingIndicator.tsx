import React, { useEffect, useState } from 'react'

/** Some locales already end the label with an ellipsis — normalize to exactly one. */
function withEllipsis(label: string): string {
  return `${label.replace(/(?:…|\.{3})+$/u, '')}…`
}

/** Pre-stream thinking indicator , two mutually exclusive phases:
 *  grow/shrink blue dots first, then after ~1.2s only the shimmer label with a " · Ns" elapsed hint after 3s.
 *  The .ai-typing-* style classes are provided by each app's styles.css. */
export function AiTypingIndicator({ label }: { readonly label: string }): React.JSX.Element {
  const [elapsed, setElapsed] = useState(0)
  const [showLabel, setShowLabel] = useState(false)

  useEffect(() => {
    const phase = setTimeout(() => setShowLabel(true), 1200)
    const timer = setInterval(() => setElapsed((s) => s + 1), 1000)
    return () => {
      clearTimeout(phase)
      clearInterval(timer)
    }
  }, [])

  return (
    <span className="ai-typing" role="status" aria-label={label}>
      {showLabel ? (
        <span className="ai-typing-label">
          {withEllipsis(label)}
          {elapsed >= 3 && (
            <span className="ai-typing-elapsed" aria-hidden>{` · ${elapsed}s`}</span>
          )}
        </span>
      ) : (
        <span className="ai-typing-dots" aria-hidden>
          <span className="ai-typing-dot-slot">
            <span className="ai-typing-dot-grow" />
          </span>
          <span className="ai-typing-dot-slot">
            <span className="ai-typing-dot-shrink" />
          </span>
        </span>
      )}
    </span>
  )
}
