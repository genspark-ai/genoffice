/**
 * Per-sender remote-image allowlist (Lane R owns the reader layer).
 * Explicit user action only ("Always allow this sender"); stored as
 * NON-SECRET prefs in localStorage alongside the account list. Envelope
 * addresses key the list — never message content, never credentials.
 */

const STORAGE_KEY = 'genoffice.mail.imageAllowlist.v1'

function storage(): Storage | null {
  try {
    if (typeof localStorage !== 'undefined') return localStorage
  } catch {
    return null
  }
  return null
}

const memoryFallback = new Set<string>()

/** Bare addr-spec, lowercased — the allowlist key for a sender. */
export function normalizeSender(raw: string): string {
  const match = raw.match(/<([^<>]+)>/)
  return (match ? match[1] : raw).trim().toLowerCase()
}

function readStored(): Set<string> {
  const store = storage()
  if (!store) return new Set(memoryFallback)
  try {
    const raw = store.getItem(STORAGE_KEY)
    if (!raw) return new Set()
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return new Set()
    return new Set(parsed.filter((item): item is string => typeof item === 'string'))
  } catch {
    return new Set()
  }
}

function writeStored(entries: Set<string>): void {
  memoryFallback.clear()
  for (const entry of entries) memoryFallback.add(entry)
  const store = storage()
  if (!store) return
  try {
    store.setItem(STORAGE_KEY, JSON.stringify([...entries].sort()))
  } catch {
    // Quota or access errors keep the in-memory copy only.
  }
}

export function loadImageAllowlist(): string[] {
  return [...readStored()].sort()
}

export function isSenderAllowed(raw: string): boolean {
  const key = normalizeSender(raw)
  if (!key) return false
  return readStored().has(key)
}

export function allowSender(raw: string): void {
  const key = normalizeSender(raw)
  if (!key) return
  const entries = readStored()
  entries.add(key)
  writeStored(entries)
}

export function revokeSender(raw: string): void {
  const key = normalizeSender(raw)
  if (!key) return
  const entries = readStored()
  entries.delete(key)
  writeStored(entries)
}
