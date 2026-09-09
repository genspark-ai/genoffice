import { useCallback, useEffect, useRef, useState } from 'react'
import { PreviewFrame, type PreviewFrameHandle } from './preview/PreviewFrame'
import type { FromInspector } from './preview/inspector-protocol'

/**
 * Present → New tab: the owner view's live preview and nothing else. The frame
 * carries the owner's inspector, which is told to behave as the published page.
 */
export function PresentView({ title }: { title: string }) {
  const [url, setUrl] = useState<string | null>(null)
  const frameRef = useRef<PreviewFrameHandle>(null)

  useEffect(() => {
    if (title) document.title = title
    let cancelled = false
    void window.htmlApi.getPreviewInfo().then((info) => {
      if (!cancelled) setUrl(info.url)
    })
    return () => {
      cancelled = true
    }
  }, [title])

  const onMessage = useCallback((msg: FromInspector) => {
    if (msg.type === 'gx:ready') frameRef.current?.post({ type: 'gx:setMode', mode: 'browse' })
  }, [])

  return (
    <div className="present-view">
      <PreviewFrame ref={frameRef} url={url} nonce={0} zoom={100} onMessage={onMessage} />
    </div>
  )
}
