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

/** a burst past this many unclaimed handoffs drops the oldest one */
const MAX_PENDING_HANDOFFS = 8
/** a burst past this many buffered bytes drops the oldest handoffs */
const MAX_PENDING_HANDOFF_BYTES = 256 * 1024 * 1024

const pending = new Map<string, { bytes: Buffer; timer: ReturnType<typeof setTimeout> }>()
let pendingBytes = 0

/** release one handoff's timer, entry and bytes; the timer must go with the entry */
function dropPending(token: string): void {
  const entry = pending.get(token)
  if (!entry) return
  clearTimeout(entry.timer)
  pending.delete(token)
  pendingBytes -= entry.bytes.length
}

/** oldest-first eviction until `bytes` fits under both caps (Map is insertion-ordered) */
function evictFor(bytes: Buffer): void {
  while (pending.size + 1 > MAX_PENDING_HANDOFFS || pendingBytes + bytes.length > MAX_PENDING_HANDOFF_BYTES) {
    const oldest = pending.keys().next()
    if (oldest.done) return
    dropPending(oldest.value)
  }
}

export function handOffBytes(bytes: Buffer): string {
  evictFor(bytes)
  const token = randomUUID()
  const timer = setTimeout(() => dropPending(token), HANDOFF_TTL_MS)
  timer.unref?.()
  pending.set(token, { bytes, timer })
  pendingBytes += bytes.length
  return `${LAZY_MEDIA_SCHEME}://${HANDOFF_HOST}/${token}`
}

/** the bytes behind a handoff URL, released on first take; null for any other URL */
export function takeHandoff(url: string): Buffer | null {
  const prefix = `${LAZY_MEDIA_SCHEME}://${HANDOFF_HOST}/`
  if (!url.startsWith(prefix)) return null
  const token = url.slice(prefix.length)
  const entry = pending.get(token)
  if (!entry) return null
  dropPending(token)
  return entry.bytes
}

export function pendingHandoffCount(): number {
  return pending.size
}
