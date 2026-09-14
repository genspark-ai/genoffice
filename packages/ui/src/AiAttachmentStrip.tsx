import React from 'react'

import { IconAttachment } from './icons'
import type { ChatAttachment } from '@genoffice/chat-runtime/types'

export interface AiAttachmentStripProps {
  attachments: ChatAttachment[]
  className?: string
}

/**
 * Read-only chip strip rendered above the chat timeline so the user can
 * see at a glance which files / images were attached to the message.
 * No add/remove UI lives here — the composer owns that.
 */
export function AiAttachmentStrip({
  attachments,
  className,
}: AiAttachmentStripProps): React.JSX.Element | null {
  if (!attachments.length) return null
  return (
    <div
      className={['ai-attachment-strip', className].filter(Boolean).join(' ')}
      aria-label={`${attachments.length} attachment${attachments.length === 1 ? '' : 's'}`}
    >
      {attachments.map(att => (
        <span key={att.id} className="ai-attachment-chip" title={att.name}>
          <IconAttachment size={12} />
          <span className="ai-attachment-chip-name">{att.name}</span>
          <span className="ai-attachment-chip-size">{formatBytes(att.size)}</span>
        </span>
      ))}
    </div>
  )
}

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return ''
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`
}
