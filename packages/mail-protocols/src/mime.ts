/**
 * MIME to view-model: parse raw RFC 5322 bytes with postal-mime into the
 * envelope/body shapes the renderer consumes. Parsing is data-only; hostile
 * HTML (scripts, remote images) is preserved verbatim for the quarantined
 * reader layer, never executed here. Helpers flag remote references and
 * build single-line snippets for the list view.
 */

import PostalMime, { type Address } from 'postal-mime'
import { decodeDisplayName, type MailAddress } from './addresses'

export interface ParsedAttachment {
  filename: string
  mimeType: string
  disposition: 'attachment' | 'inline' | null
  contentId: string | undefined
  size: number
  content: Uint8Array
}

export interface ParsedMail {
  messageId: string | undefined
  from: MailAddress[]
  to: MailAddress[]
  cc: MailAddress[]
  subject: string
  date: string | undefined
  text: string
  html: string | undefined
  snippet: string
  attachments: ParsedAttachment[]
  hasAttach: boolean
  /** True when the HTML body references remote resources (blocked by default). */
  hasRemoteRefs: boolean
}

const REMOTE_REF_RE = /(<[^>]+\s(?:src|href|poster|background)\s*=\s*["']?\s*https?:)/i

function toAddresses(list: Address[] | Address | undefined): MailAddress[] {
  const arr = list === undefined ? [] : Array.isArray(list) ? list : [list]
  const flat: MailAddress[] = []
  for (const a of arr) {
    if (a.group) {
      for (const m of a.group)
        flat.push({ name: decodeDisplayName(m.name ?? ''), address: m.address ?? '' })
    } else if (a.address) {
      flat.push({ name: decodeDisplayName(a.name ?? ''), address: a.address })
    }
  }
  return flat
}

/** Normalize attachment bytes to a Uint8Array view. */
function toBytes(content: ArrayBuffer | Uint8Array | string): Uint8Array {
  if (typeof content === 'string') return new TextEncoder().encode(content)
  if (content instanceof Uint8Array) return content
  return new Uint8Array(content)
}

/** Fold a header value to one display line (kills CRLF injection). */
export function singleLine(value: string | undefined): string {
  return (value ?? '').replace(/[\r\n]+/g, ' ').trim()
}

/** Build a list-row snippet: whitespace-folded, capped, no newlines. */
export function makeSnippet(text: string, max = 200): string {
  const folded = text.replace(/\s+/g, ' ').trim()
  return folded.length > max ? `${folded.slice(0, max - 1).trimEnd()}…` : folded
}

/** True when HTML references remote http(s) resources. */
export function detectRemoteRefs(html: string | undefined): boolean {
  return !!html && REMOTE_REF_RE.test(html)
}

/**
 * Parse a raw message into the view-model. Accepts string or bytes.
 * Rejects only on malformed input that postal-mime cannot scan.
 */
export async function parseMimeMessage(raw: string | Uint8Array): Promise<ParsedMail> {
  const email = await PostalMime.parse(raw)
  const text = email.text ?? ''
  const html = email.html ?? undefined
  const attachments: ParsedAttachment[] = (email.attachments ?? []).map((a) => {
    const content = toBytes(a.content ?? new Uint8Array(0))
    return {
      filename: a.filename ?? 'unnamed',
      mimeType: a.mimeType ?? 'application/octet-stream',
      disposition: (a.disposition ?? null) as ParsedAttachment['disposition'],
      contentId: a.contentId ?? undefined,
      size: content.byteLength,
      content,
    }
  })
  return {
    messageId: email.messageId,
    from: toAddresses(email.from ? [email.from] : []),
    to: toAddresses(email.to),
    cc: toAddresses(email.cc),
    subject: singleLine(email.subject),
    date: email.date,
    text,
    html,
    snippet: makeSnippet(text || stripTags(html ?? '')),
    attachments,
    hasAttach: attachments.length > 0,
    hasRemoteRefs: detectRemoteRefs(html),
  }
}

/** Crude tag strip for snippet fallback when no text part exists. */
export function stripTags(html: string): string {
  return html
    .replace(/<script[\s\S]*?<\/script\s*>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style\s*>/gi, ' ')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
}
