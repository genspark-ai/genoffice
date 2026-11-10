/**
 * Renderer side of Word's face resolution: the generated word-faces.json plus
 * the main process's installed-font index (IPC docs:font-metrics), prefetched
 * at document open so every later call is synchronous.
 */
import type { FaceVerticalMetrics } from '@genoffice/font-metrics'
import {
  faceKey,
  resolveFace,
  storeContext,
  type FaceHints,
  type FaceMetrics,
  type ResolvedFace,
  type StoredFace,
  type WordFacesData,
} from '@genoffice/font-metrics/resolve'
import wordFaces from '@genoffice/font-metrics/word-faces.json'

const EA_CODE_PAGES = 0x3e0000 // ulCodePageRange1 bits 17-21

/** the IPC rejected: neither installed nor the snapshot, the store/substitution chain answers */
const UNKNOWN = Symbol('unknown')
/** IPC answers; null = not installed, absent = not asked yet */
const installed = new Map<string, StoredFace | null | typeof UNKNOWN>()
const inflight = new Map<string, Promise<void>>()
let provider: ((family: string) => Promise<FaceVerticalMetrics | null>) | null | undefined
let hints = new Map<string, FaceHints>()
let embedded = new Map<string, FaceMetrics>()
const resolved = new Map<string, ResolvedFace | null>()

function fromVertical(key: string, m: FaceVerticalMetrics): StoredFace {
  const u = m.unitsPerEm
  return {
    family: key,
    source: 'installed',
    hheaTotal: (m.hheaAscender - m.hheaDescender + m.hheaLineGap) / u,
    ascent: m.hheaAscender / u,
    descent: -m.hheaDescender / u,
    xAvg: m.xAvgCharWidth == null ? null : m.xAvgCharWidth / u,
    eastAsian: ((m.codePageRange1 ?? 0) & EA_CODE_PAGES) !== 0,
  }
}

/** test seam: replaces window.desktop.fontMetrics; null disables the index */
export function setFontMetricsProvider(
  fn: ((family: string) => Promise<FaceVerticalMetrics | null>) | null | undefined,
): void {
  provider = fn
  installed.clear()
  inflight.clear()
  resolved.clear()
}

function metricsIpc(): ((family: string) => Promise<FaceVerticalMetrics | null>) | null {
  if (provider !== undefined) return provider
  if (typeof window === 'undefined') return null
  const fn = window.desktop?.fontMetrics
  return typeof fn === 'function' ? fn : null
}

/** One IPC round per unseen family; resolves once every answer is cached */
export function prefetchFaces(names: Iterable<string>): Promise<void> {
  const ipc = metricsIpc()
  if (!ipc) return Promise.resolve()
  const waits: Promise<void>[] = []
  for (const name of names) {
    const key = faceKey(name)
    if (!key || installed.has(key)) continue
    let p = inflight.get(key)
    if (!p) {
      p = ipc(key)
        .then(
          (m) => installed.set(key, m ? fromVertical(key, m) : null),
          () => installed.set(key, UNKNOWN),
        )
        .then(() => {
          inflight.delete(key)
          invalidate(key)
        })
      inflight.set(key, p)
    }
    waits.push(p)
  }
  return Promise.all(waits).then(() => undefined)
}

/** drops the answers whose chain touched the family (null answers keep no path) */
function invalidate(key: string): void {
  for (const [name, face] of resolved) {
    if (!face || face.path.includes(key)) resolved.delete(name)
  }
}

const ctx = storeContext(wordFaces as WordFacesData, {
  installed: (key) => {
    if (!metricsIpc()) return undefined
    const hit = installed.get(key)
    if (hit === undefined) void prefetchFaces([key])
    return hit === UNKNOWN ? null : hit
  },
  hints: (key) => hints.get(key),
  embedded: (key) => embedded.get(key) ?? null,
})

export function setResolverFontTable(
  entries: ReadonlyArray<{ name: string } & FaceHints> | null | undefined,
): void {
  hints = new Map()
  for (const e of entries ?? []) {
    const { name, ...rest } = e
    hints.set(faceKey(name), rest)
  }
  resolved.clear()
}

/** first entry per family wins */
export function setResolverEmbeddedFaces(
  faces: ReadonlyArray<{
    family: string
    ascent: number
    descent: number
    lineGap: number
    eastAsian: boolean
  }>,
): void {
  embedded = new Map()
  for (const f of faces) {
    const key = faceKey(f.family)
    const total = f.ascent + f.descent + f.lineGap
    if (embedded.has(key) || !(total > 0.5 && total < 4)) continue
    embedded.set(key, {
      hheaTotal: total,
      ascent: f.ascent,
      descent: f.descent,
      xAvg: null,
      eastAsian: f.eastAsian,
    })
  }
  resolved.clear()
}

/** Word's layout face for a declared family; null while its metrics are unknown */
export function resolveDeclaredFace(name: string): ResolvedFace | null {
  const key = faceKey(name)
  const cached = resolved.get(key)
  if (cached !== undefined) return cached
  const out = resolveFace(name, ctx)
  resolved.set(key, out)
  return out
}

export interface FontResolveDiff {
  name: string
  legacy: number
  resolver: number | null
  face: string | null
  source: string | null
}

declare global {
  var __fontResolveDiff: FontResolveDiff[] | undefined
}

const reported = new Set<string>()

/** Shadow mode: the old table answers, disagreements are recorded for the cutover audit */
export function recordResolveDiff(name: string, legacy: number, face: ResolvedFace | null): void {
  const key = faceKey(name)
  if (reported.has(key)) return
  reported.add(key)
  const diff: FontResolveDiff = {
    name,
    legacy,
    resolver: face?.lineFactor ?? null,
    face: face?.family ?? null,
    source: face?.source ?? null,
  }
  ;(globalThis.__fontResolveDiff ??= []).push(diff)
  if (import.meta.env?.DEV) console.warn('[font-resolve] shadow disagreement', diff)
}
