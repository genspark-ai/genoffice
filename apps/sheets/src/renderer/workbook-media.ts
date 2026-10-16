/**
 * Fetches a workbook image as a data URL. Lives apart from the visual layer
 * so a picture inside a note can be shown the same way an in-cell picture is,
 * by id, without either place owning the other's component.
 */
import { useEffect, useState } from 'react'

import { isMetafileMime, metafileToDataUrl } from '@genoffice/docx-engine/metafile'

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
  return bytes
}

/// Lazily fetches a workbook image as a data URL (metafiles rasterize to PNG
/// first). Pass an undefined sessionId to skip fetching.
export function useWorkbookMediaUrl(
  sessionId: string | undefined,
  visualId: string,
): { url: string | null; unavailable: boolean } {
  const [url, setUrl] = useState<string | null>(null)
  const [unavailable, setUnavailable] = useState(false)

  useEffect(() => {
    if (sessionId === undefined) return
    let isCurrent = true
    void window.desktopApi
      .readWorkbookMedia({ sessionId, visualId })
      .then(async (media) => {
        const next = isMetafileMime(media.mediaType)
          ? await metafileToDataUrl(base64ToBytes(media.base64), media.mediaType)
          : `data:${media.mediaType};base64,${media.base64}`
        if (!isCurrent) return
        if (next) setUrl(next)
        else setUnavailable(true)
      })
      .catch((reason: unknown) => {
        // Unsupported or unreadable media degrades to an empty frame; the
        // grid must never show the failure text.
        console.warn(`workbook media unavailable (${visualId})`, reason)
        if (isCurrent) setUnavailable(true)
      })
    return () => {
      isCurrent = false
    }
  }, [sessionId, visualId])

  return { url, unavailable }
}
