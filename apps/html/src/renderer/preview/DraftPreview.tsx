import { useEffect, useRef, type ReactElement } from 'react'

/**
 * Read-only mirror of a page while the AI is still writing it. The draft grows
 * by appending, so it is streamed into one same-origin frame with
 * document.write: the parser picks up where it left off and nothing reloads.
 * (A srcdoc frame is cross-origin and reloads on every change; Chromium then
 * withholds its first paint after each swap, which read as a white mirror.)
 * Scripts stay off through the sandbox; this is a progress view, not the document.
 */
export function DraftPreview({ html }: { html: string }): ReactElement {
  const frameRef = useRef<HTMLIFrameElement>(null)
  const writtenRef = useRef('')

  useEffect(() => {
    const doc = frameRef.current?.contentDocument
    if (!doc) return
    const written = writtenRef.current
    if (written && html.startsWith(written)) {
      if (html.length > written.length) doc.write(html.slice(written.length))
    } else {
      doc.open()
      doc.write(html)
    }
    writtenRef.current = html
  }, [html])

  return (
    <div className="draft-preview" aria-hidden>
      <iframe
        ref={frameRef}
        className="draft-preview-frame"
        title="draft"
        sandbox="allow-same-origin"
      />
    </div>
  )
}
