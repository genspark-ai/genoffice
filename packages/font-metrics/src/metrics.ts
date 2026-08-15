/**
 * Vertical metrics of installed fonts, resolved by exact family name (nameID
 * 1/16, all languages, normalized) — never fuzzy: a wrong-font match is worse
 * than a miss, the caller's heuristic fallback handles misses.
 *
 * Both metric groups are returned: Word uses win metrics unless OS/2
 * fsSelection bit 7 (USE_TYPO_METRICS) is set, in which case it honors the
 * typo group — deciding is the consumer's policy, not this package's.
 */
import { closeSync, mkdirSync, openSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'

import { getFontIndex, norm, readTable, readTableDir, styleScore } from './sfnt'

export interface FaceVerticalMetrics {
  unitsPerEm: number
  hheaAscender: number
  hheaDescender: number
  hheaLineGap: number
  typoAscender: number | null
  typoDescender: number | null
  typoLineGap: number | null
  winAscent: number | null
  winDescent: number | null
  /** OS/2 fsSelection bit 7 (USE_TYPO_METRICS) */
  useTypoMetrics: boolean
}

function parseFaceMetrics(fd: number, offset: number): FaceVerticalMetrics | null {
  const tables = readTableDir(fd, offset)
  if (!tables) return null
  // head 54B / hhea 36B / OS/2 ≤ ~100B: bounded targeted reads, like the name index
  const head = readTable(fd, tables, 'head', 512)
  const hhea = readTable(fd, tables, 'hhea', 512)
  if (!head || head.length < 20 || !hhea || hhea.length < 10) return null
  const unitsPerEm = head.readUInt16BE(18)
  if (unitsPerEm === 0) return null
  const os2 = readTable(fd, tables, 'OS/2', 512)
  const hasTypo = !!os2 && os2.length >= 74
  const hasWin = !!os2 && os2.length >= 78
  return {
    unitsPerEm,
    hheaAscender: hhea.readInt16BE(4),
    hheaDescender: hhea.readInt16BE(6),
    hheaLineGap: hhea.readInt16BE(8),
    typoAscender: hasTypo ? os2.readInt16BE(68) : null,
    typoDescender: hasTypo ? os2.readInt16BE(70) : null,
    typoLineGap: hasTypo ? os2.readInt16BE(72) : null,
    winAscent: hasWin ? os2.readUInt16BE(74) : null,
    winDescent: hasWin ? os2.readUInt16BE(76) : null,
    useTypoMetrics: !!os2 && os2.length >= 64 && (os2.readUInt16BE(62) & 0x80) !== 0,
  }
}

interface CacheEntry {
  path: string
  offset: number
  mtimeMs: number
  size: number
  /** null = family known to be uninstalled at (re)build time; revalidated per process */
  m: FaceVerticalMetrics | null
}

let cacheFile: string | null = null
let persisted: Record<string, CacheEntry> | null = null
const misses = new Set<string>()

/** Point the persistent cache at a directory (e.g. Electron userData); optional. */
export function configureMetricsCache(dir: string): void {
  cacheFile = join(dir, 'font-vertical-metrics.json')
  persisted = null
  misses.clear()
}

function loadCache(): Record<string, CacheEntry> {
  if (persisted) return persisted
  persisted = {}
  if (cacheFile) {
    try {
      const parsed: unknown = JSON.parse(readFileSync(cacheFile, 'utf8'))
      if (parsed && typeof parsed === 'object') persisted = parsed as Record<string, CacheEntry>
    } catch {
      /* absent or corrupt cache: rebuild lazily */
    }
  }
  return persisted
}

function saveCache(): void {
  if (!cacheFile || !persisted) return
  try {
    mkdirSync(dirname(cacheFile), { recursive: true })
    writeFileSync(cacheFile, JSON.stringify(persisted))
  } catch {
    /* cache is an optimization; queries still work without it */
  }
}

function cacheValid(entry: CacheEntry): boolean {
  try {
    const st = statSync(entry.path)
    return st.mtimeMs === entry.mtimeMs && st.size === entry.size
  } catch {
    return false
  }
}

/**
 * Vertical metrics of the installed family's Regular-ish face, or null when the
 * family is not installed (exact normalized-name match only). Results persist
 * across runs keyed by font path+mtime+size; installed-set changes invalidate
 * naturally (missing families are only cached per process).
 */
export function familyVerticalMetrics(family: string): FaceVerticalMetrics | null {
  const key = norm(family)
  if (!key) return null
  if (misses.has(key)) return null
  const cache = loadCache()
  const hit = cache[key]
  if (hit && cacheValid(hit)) return hit.m

  const candidates = getFontIndex().byFamily.get(key)
  const face = candidates?.length
    ? [...candidates].sort((a, b) => styleScore(b, []) - styleScore(a, []))[0]
    : undefined
  if (!face) {
    misses.add(key)
    if (hit) {
      delete cache[key]
      saveCache()
    }
    return null
  }

  let metrics: FaceVerticalMetrics | null = null
  let st: { mtimeMs: number; size: number } | null = null
  let parsed = false
  try {
    st = statSync(face.path)
    const fd = openSync(face.path, 'r')
    try {
      // a null here is deterministic (non-sfnt face / missing tables) and safe
      // to persist; thrown reads are transient (Windows font locks) and are not
      metrics = parseFaceMetrics(fd, face.offset)
      parsed = true
    } finally {
      closeSync(fd)
    }
  } catch {
    /* transient stat/open/read failure: no cache write, retried next call */
  }
  if (st && parsed) {
    cache[key] = {
      path: face.path,
      offset: face.offset,
      mtimeMs: st.mtimeMs,
      size: st.size,
      m: metrics,
    }
    saveCache()
  }
  return metrics
}
