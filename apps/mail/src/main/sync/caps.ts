import {
  GMAIL_MAX_MESSAGE_BYTES,
  GMAIL_MAX_RECIPIENTS_PER_MESSAGE,
  GMAIL_WARN_BYTES,
  GRAPH_MAX_ATTACHMENT_BYTES,
  GRAPH_MAX_RECIPIENTS,
  GRAPH_MAX_TOTAL_BYTES,
  GRAPH_WARN_TOTAL_BYTES,
  OUTGOING_WARN_BYTES,
  type SendRequest,
} from '@genoffice/mail-protocols'

/**
 * Per-driver send caps enforced by the sync engine (P4 spec section 3).
 * Caps run twice: at enqueue for a fast reject, and at send in case caps
 * changed while the entry waited in the outbox.
 */

export type AccountKind = 'imap' | 'gmail' | 'graph'

export interface SendCapSet {
  warnTotalBytes: number
  maxTotalBytes: number
  maxAttachmentBytes: number
  maxRecipients: number
}

/** Locked per-driver caps (P1 values, mirrored here for enqueue checks). */
export function capsForKind(kind: AccountKind): SendCapSet {
  switch (kind) {
    case 'gmail':
      return {
        warnTotalBytes: GMAIL_WARN_BYTES,
        maxTotalBytes: GMAIL_MAX_MESSAGE_BYTES,
        // Gmail caps the whole encoded message; no per-file ceiling.
        maxAttachmentBytes: GMAIL_MAX_MESSAGE_BYTES,
        maxRecipients: GMAIL_MAX_RECIPIENTS_PER_MESSAGE,
      }
    case 'graph':
      return {
        warnTotalBytes: GRAPH_WARN_TOTAL_BYTES,
        maxTotalBytes: GRAPH_MAX_TOTAL_BYTES,
        maxAttachmentBytes: GRAPH_MAX_ATTACHMENT_BYTES,
        maxRecipients: GRAPH_MAX_RECIPIENTS,
      }
    case 'imap':
    default:
      return {
        warnTotalBytes: OUTGOING_WARN_BYTES,
        // Generic IMAP/SMTP ceilings (P1 ImapSmtpDriver values).
        maxTotalBytes: 50 * 1024 * 1024,
        maxAttachmentBytes: 25 * 1024 * 1024,
        maxRecipients: 100,
      }
  }
}

/**
 * Rough encoded size of a send request: UTF-8 body bytes times base64
 * overhead plus staged attachment bytes. Mirrors the Gmail estimator so
 * enqueue verdicts match driver verdicts for the same request.
 */
export function estimateOutboxBytes(req: SendRequest): number {
  const textBytes = Buffer.byteLength(req.text, 'utf8') + Buffer.byteLength(req.html ?? '', 'utf8')
  const attachBytes = (req.attachments ?? []).reduce((sum, a) => sum + a.size, 0)
  return Math.ceil(((textBytes + attachBytes) * 4) / 3) + 4096
}

export function countOutboxRecipients(req: SendRequest): number {
  return req.to.length + (req.cc?.length ?? 0) + (req.bcc?.length ?? 0)
}

export type CapVerdict =
  'ok' | 'warn' | 'too-large' | 'too-many-recipients' | 'attachment-too-large'

export interface CapCheck {
  verdict: CapVerdict
  /** Human-readable, PII-free reason for rejections (safe for lastError). */
  message?: string
}

/**
 * Enforce per-driver send limits. Returns ok/warn, or the rejection
 * reason when the request must not be queued or transmitted.
 */
export function checkOutboxCaps(
  kind: AccountKind,
  req: SendRequest,
  estimatedBytes?: number,
): CapCheck {
  const caps = capsForKind(kind)
  if (countOutboxRecipients(req) > caps.maxRecipients) {
    return {
      verdict: 'too-many-recipients',
      message: `recipient count above the per-message cap (${caps.maxRecipients})`,
    }
  }
  for (const att of req.attachments ?? []) {
    if (att.size > caps.maxAttachmentBytes) {
      return {
        verdict: 'attachment-too-large',
        message: `attachment above the per-file cap (${caps.maxAttachmentBytes} bytes)`,
      }
    }
  }
  const bytes = estimatedBytes ?? estimateOutboxBytes(req)
  if (bytes > caps.maxTotalBytes) {
    return {
      verdict: 'too-large',
      message: `encoded size above the per-message cap (${caps.maxTotalBytes} bytes)`,
    }
  }
  if (bytes > caps.warnTotalBytes) return { verdict: 'warn' }
  return { verdict: 'ok' }
}
