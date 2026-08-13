import { randomUUID } from 'node:crypto'
import { mkdir, readFile, rename, writeFile } from 'node:fs/promises'
import { dirname } from 'node:path'
import type { SavedSignature, SignatureData } from '../shared/ipc'

/** Oldest entries are evicted beyond this; keeps the userData JSON bounded (images are base64 PNGs) */
export const MAX_SAVED_SIGNATURES = 10

const isFiniteDim = (v: unknown): v is number =>
  typeof v === 'number' && Number.isFinite(v) && v > 0

/** Runtime validation for IPC payloads and for entries read back from disk */
export function isSignatureData(value: unknown): value is SignatureData {
  if (typeof value !== 'object' || value === null) return false
  const v = value as Record<string, unknown>
  if (!isFiniteDim(v.width) || !isFiniteDim(v.height)) return false
  if (v.kind === 'image') return typeof v.image === 'string' && v.image.length > 0
  if (v.kind === 'strokes') {
    return (
      Array.isArray(v.paths) &&
      v.paths.length > 0 &&
      v.paths.every(
        (p) =>
          Array.isArray(p) &&
          p.length >= 4 &&
          p.every((n) => typeof n === 'number' && Number.isFinite(n)),
      )
    )
  }
  return false
}

/** Drop malformed entries (hand-edited file, older formats) instead of failing the whole list */
export function sanitizeSignatures(raw: unknown): SavedSignature[] {
  if (!Array.isArray(raw)) return []
  const out: SavedSignature[] = []
  for (const entry of raw) {
    if (typeof entry !== 'object' || entry === null) continue
    const e = entry as Record<string, unknown>
    if (typeof e.id !== 'string' || e.id.length === 0) continue
    if (!isSignatureData(e.data)) continue
    if (out.some((s) => s.id === e.id)) continue
    const createdAt =
      typeof e.createdAt === 'number' && Number.isFinite(e.createdAt) ? e.createdAt : 0
    out.push({ id: e.id, createdAt, data: e.data })
  }
  return out
}

/** Same payload = same signature: re-saving refreshes its position instead of duplicating */
const sameData = (a: SignatureData, b: SignatureData): boolean => {
  if (a.kind !== b.kind || a.width !== b.width || a.height !== b.height) return false
  if (a.kind === 'image' && b.kind === 'image') return a.image === b.image
  if (a.kind === 'strokes' && b.kind === 'strokes') {
    return (
      a.paths.length === b.paths.length &&
      a.paths.every((p, i) => {
        const q = b.paths[i]!
        return p.length === q.length && p.every((n, j) => n === q[j])
      })
    )
  }
  return false
}

/** Prepend a signature (newest first), dedupe identical payloads, cap the list size */
export function addSignature(
  list: SavedSignature[],
  data: SignatureData,
  now: number = Date.now(),
): SavedSignature[] {
  const rest = list.filter((s) => !sameData(s.data, data))
  return [{ id: randomUUID(), createdAt: now, data }, ...rest].slice(0, MAX_SAVED_SIGNATURES)
}

export function removeSignature(list: SavedSignature[], id: string): SavedSignature[] {
  return list.filter((s) => s.id !== id)
}

/** Missing or unreadable file is just an empty list; a broken JSON must not break signing */
export async function loadSignatures(filePath: string): Promise<SavedSignature[]> {
  try {
    return sanitizeSignatures(JSON.parse(await readFile(filePath, 'utf8')))
  } catch {
    return []
  }
}

/** Atomic write (tmp + rename) so a crash mid-write can't corrupt the saved list */
export async function saveSignatures(filePath: string, list: SavedSignature[]): Promise<void> {
  await mkdir(dirname(filePath), { recursive: true })
  const tmp = `${filePath}.${process.pid}.tmp`
  await writeFile(tmp, JSON.stringify(list))
  await rename(tmp, filePath)
}
