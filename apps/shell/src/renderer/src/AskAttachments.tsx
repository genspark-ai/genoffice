import { useEffect, useState } from 'react'
import { ATTACHMENT_IMAGE_EXTS, formatKbSize, type AttachmentInfo } from '@genoffice/agent-core'
import { FileBadge } from './file-badge'

/**
 * The files attached to an Ask AI question: in the composer header (with a
 * remove button) and echoed read-only on the sent message. Images show a
 * thumbnail read through the attachment IPC; other files show a type badge.
 */
export function AskAttachments({
  items,
  onRemove,
  removeLabel,
}: {
  items: readonly AttachmentInfo[]
  onRemove?: (path: string) => void
  removeLabel?: string
}): React.JSX.Element | null {
  const [previews, setPreviews] = useState<Record<string, string>>({})
  const imagePaths = items
    .filter((a) => ATTACHMENT_IMAGE_EXTS.has(a.ext))
    .map((a) => a.path)
    .join('\0')
  useEffect(() => {
    let live = true
    for (const path of imagePaths ? imagePaths.split('\0') : []) {
      if (previews[path]) continue
      void window.aiOffice
        .readAttachmentImage(path)
        .then((r) => {
          if (live && r.ok && r.base64 && r.mime)
            setPreviews((prev) => ({ ...prev, [path]: `data:${r.mime};base64,${r.base64}` }))
        })
        .catch(() => undefined)
    }
    return () => {
      live = false
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps -- keyed on the image set only
  }, [imagePaths])
  if (items.length === 0) return null
  return (
    <div className={`kb-ask-attachments${onRemove ? '' : ' kb-ask-attachments-sent'}`}>
      {items.map((a) => (
        <span key={a.path} className="kb-ask-attachment" title={a.path}>
          {previews[a.path] ? (
            <img className="kb-ask-attachment-thumb" src={previews[a.path]} alt={a.name} />
          ) : (
            <span className="kb-ask-attachment-icon">
              <FileBadge ext={a.ext} size={20} />
            </span>
          )}
          <span className="kb-ask-attachment-meta">
            <span className="kb-ask-attachment-name">{a.name}</span>
            <span className="kb-ask-attachment-size">{formatKbSize(a.sizeBytes)}</span>
          </span>
          {onRemove && (
            <button
              type="button"
              className="kb-ask-attachment-remove"
              onClick={() => onRemove(a.path)}
              aria-label={removeLabel}
              data-tip={removeLabel}
            >
              <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden>
                <path
                  d="M2 2l6 6M8 2l-6 6"
                  stroke="currentColor"
                  strokeWidth="1.5"
                  strokeLinecap="round"
                />
              </svg>
            </button>
          )}
        </span>
      ))}
    </div>
  )
}
