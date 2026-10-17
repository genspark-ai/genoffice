/**
 * Mail tab identity: renderer state carries a `mailRef` pseudo-target
 * (`mail://<accountId>/<uid>`); the shared tab record never holds a
 * `mail://` URL (locked D3). Pure helpers, no Electron, no secrets.
 */

/** `mail://<accountId>/<uid>` — both segments are URI-encoded on format. */
const MAIL_REF_RE = /^mail:\/\/([^/]+)\/(.+)$/

export interface ParsedMailRef {
  accountId: string
  uid: string
}

export function formatMailRef(accountId: string, uid: string): string {
  return `mail://${encodeURIComponent(accountId)}/${encodeURIComponent(uid)}`
}

/** Parse a mailRef back into its segments; null when the shape is wrong. */
export function parseMailRef(ref: unknown): ParsedMailRef | null {
  if (typeof ref !== 'string') return null
  const match = MAIL_REF_RE.exec(ref.trim())
  if (!match) return null
  const accountId = decodeURIComponentSafe(match[1])
  const uid = decodeURIComponentSafe(match[2])
  if (!accountId || !uid) return null
  return { accountId, uid }
}

/** Shape guard for the `mail:open-message` payload path (router never sees mailRefs). */
export function isMailRef(ref: unknown): ref is string {
  return parseMailRef(ref) !== null
}

function decodeURIComponentSafe(segment: string): string | null {
  try {
    return decodeURIComponent(segment)
  } catch {
    return null
  }
}
