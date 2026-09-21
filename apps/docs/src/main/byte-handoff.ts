import { randomUUID } from 'node:crypto'
import { LAZY_MEDIA_SCHEME } from '@genoffice/docx-engine/lazy-media'

/**
 * Document bytes cross to the renderer over the media protocol instead of an
 * IPC reply: the reply serializer copies the whole buffer (and doubles its
 * scratch space while growing), which is what took the main process down on a
 * 1 GB file. A one-shot token URL hands the buffer to a single fetch.
 */

const HANDOFF_HOST = 'handoff'
/** an unclaimed handoff (renderer gone before its fetch) is dropped after this */
const HANDOFF_TTL_MS = 60_000

const pending = new Map<string, { bytes: Buffer; timer: ReturnType<typeof setTimeout> }>()

export function handOffBytes(bytes: Buffer): string {
  const token = randomUUID()
  const timer = setTimeout(() => pending.delete(token), HANDOFF_TTL_MS)
  timer.unref?.()
  pending.set(token, { bytes, timer })
  return `${LAZY_MEDIA_SCHEME}://${HANDOFF_HOST}/${token}`
}

/** the bytes behind a handoff URL, released on first take; null for any other URL */
export function takeHandoff(url: string): Buffer | null {
  const prefix = `${LAZY_MEDIA_SCHEME}://${HANDOFF_HOST}/`
  if (!url.startsWith(prefix)) return null
  const token = url.slice(prefix.length)
  const entry = pending.get(token)
  if (!entry) return null
  clearTimeout(entry.timer)
  pending.delete(token)
  return entry.bytes
}

export function pendingHandoffCount(): number {
  return pending.size
}
