/** One contiguous source replacement; the only primitive every edit compiles down to. */
export interface Patch {
  from: number
  to: number
  text: string
}

export type PatchOrigin = 'manual' | 'ai' | 'inspector' | 'format' | 'load'

export interface PatchSet {
  patches: Patch[]
  baseVersion: number
  origin: PatchOrigin
  label: string
}

export type PatchError =
  | { kind: 'stale'; baseVersion: number; currentVersion: number }
  | { kind: 'bounds'; index: number }
  | { kind: 'overlap'; index: number }

export function validatePatchSet(
  set: PatchSet,
  currentVersion: number,
  length: number,
): PatchError | null {
  if (set.baseVersion !== currentVersion) {
    return { kind: 'stale', baseVersion: set.baseVersion, currentVersion }
  }
  const sorted = sortPatches(set.patches)
  let cursor = -1
  for (let i = 0; i < sorted.length; i++) {
    const p = sorted[i]!
    if (p.from < 0 || p.to < p.from || p.to > length) return { kind: 'bounds', index: i }
    if (p.from < cursor) return { kind: 'overlap', index: i }
    cursor = p.to
  }
  return null
}

export function sortPatches(patches: readonly Patch[]): Patch[] {
  return [...patches].sort((a, b) => a.from - b.from || a.to - b.to)
}

/** Apply already-validated patches to text (ascending order, spliced from the end). */
export function applyPatches(text: string, patches: readonly Patch[]): string {
  const sorted = sortPatches(patches)
  let out = text
  for (let i = sorted.length - 1; i >= 0; i--) {
    const p = sorted[i]!
    out = out.slice(0, p.from) + p.text + out.slice(p.to)
  }
  return out
}

/** Inverse patches (in the coordinate space of the patched text), for tests and snapshots. */
export function invertPatches(text: string, patches: readonly Patch[]): Patch[] {
  const sorted = sortPatches(patches)
  const out: Patch[] = []
  let delta = 0
  for (const p of sorted) {
    out.push({
      from: p.from + delta,
      to: p.from + delta + p.text.length,
      text: text.slice(p.from, p.to),
    })
    delta += p.text.length - (p.to - p.from)
  }
  return out
}
