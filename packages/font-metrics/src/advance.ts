/**
 * Horizontal advance widths of installed fonts (cmap + hmtx), resolved by
 * exact family name like the vertical metrics — never fuzzy. Consumers
 * preflight rebuilt text against the OUTPUT font so a substituted face's
 * wider advances are caught before they reflow the document.
 *
 * Kerning (kern/GPOS) is deliberately ignored: plain advances bound the
 * un-kerned line width, and callers compare with their own slack. Shaping
 * (Arabic joining, complex scripts) is out of scope — callers should not
 * measure those scripts here.
 */
import { closeSync, openSync } from 'node:fs'

import { getFontIndex, norm, readTable, readTableDir, styleScore } from './sfnt'
import type { FaceRef } from './sfnt'

export interface AdvanceStyle {
  bold?: boolean
  italic?: boolean
}

/** hmtx can be numGlyphs×4B (64K glyphs → 256KB) — bounded, but far past the
 * 512B vertical-metrics reads; cmap format-12 tables on CJK fonts reach ~100KB */
const HMTX_MAX = 1 << 19
const CMAP_MAX = 1 << 20

interface ParsedFace {
  unitsPerEm: number
  /** advance width in font units per glyph id; ids past the array reuse the last entry */
  advances: Uint16Array
  /** raw cmap subtable (format 4 or 12), looked up per codepoint via binary search */
  cmapFormat: 4 | 12
  cmap: Buffer
}

function parseCmap(buf: Buffer): { format: 4 | 12; sub: Buffer } | null {
  if (buf.length < 4) return null
  const count = buf.readUInt16BE(2)
  let best: { off: number; rank: number } | null = null
  for (let i = 0; i < count; i++) {
    const r = 4 + 8 * i
    if (r + 8 > buf.length) break
    const platform = buf.readUInt16BE(r)
    const encoding = buf.readUInt16BE(r + 2)
    const off = buf.readUInt32BE(r + 4)
    // full-Unicode (3/10, 0/4+) beats BMP (3/1, 0/0-3); everything else loses
    const rank =
      platform === 3 && encoding === 10
        ? 4
        : platform === 0 && encoding >= 4
          ? 3
          : platform === 3 && encoding === 1
            ? 2
            : platform === 0
              ? 1
              : 0
    if (rank > 0 && (!best || rank > best.rank)) best = { off, rank }
  }
  if (!best || best.off + 4 > buf.length) return null
  const sub = buf.subarray(best.off)
  const format = sub.readUInt16BE(0)
  if (format === 4 || format === 12) return { format, sub }
  return null
}

/** glyph id for a codepoint, 0 (.notdef) when unmapped */
function glyphOf(face: ParsedFace, cp: number): number {
  const sub = face.cmap
  if (face.cmapFormat === 12) {
    const nGroups = sub.readUInt32BE(12)
    let lo = 0
    let hi = nGroups - 1
    while (lo <= hi) {
      const mid = (lo + hi) >> 1
      const g = 16 + 12 * mid
      if (g + 12 > sub.length) return 0
      const start = sub.readUInt32BE(g)
      const end = sub.readUInt32BE(g + 4)
      if (cp < start) hi = mid - 1
      else if (cp > end) lo = mid + 1
      else return sub.readUInt32BE(g + 8) + (cp - start)
    }
    return 0
  }
  // format 4: BMP only
  if (cp > 0xffff) return 0
  const segCountX2 = sub.readUInt16BE(6)
  const endBase = 14
  const startBase = endBase + segCountX2 + 2
  const deltaBase = startBase + segCountX2
  const rangeBase = deltaBase + segCountX2
  let lo = 0
  let hi = segCountX2 / 2 - 1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    if (endBase + 2 * mid + 2 > sub.length) return 0
    const end = sub.readUInt16BE(endBase + 2 * mid)
    if (cp > end) {
      lo = mid + 1
      continue
    }
    const start = sub.readUInt16BE(startBase + 2 * mid)
    if (cp < start) {
      hi = mid - 1
      continue
    }
    const delta = sub.readInt16BE(deltaBase + 2 * mid)
    const rangeOff = sub.readUInt16BE(rangeBase + 2 * mid)
    if (rangeOff === 0) return (cp + delta) & 0xffff
    const gi = rangeBase + 2 * mid + rangeOff + 2 * (cp - start)
    if (gi + 2 > sub.length) return 0
    const glyph = sub.readUInt16BE(gi)
    return glyph === 0 ? 0 : (glyph + delta) & 0xffff
  }
  return 0
}

function parseFaceAdvances(fd: number, offset: number): ParsedFace | null {
  const tables = readTableDir(fd, offset)
  if (!tables) return null
  const head = readTable(fd, tables, 'head', 512)
  const hhea = readTable(fd, tables, 'hhea', 512)
  if (!head || head.length < 20 || !hhea || hhea.length < 36) return null
  const unitsPerEm = head.readUInt16BE(18)
  const numberOfHMetrics = hhea.readUInt16BE(34)
  if (unitsPerEm === 0 || numberOfHMetrics === 0) return null
  const hmtx = readTable(fd, tables, 'hmtx', HMTX_MAX)
  const cmapRaw = readTable(fd, tables, 'cmap', CMAP_MAX)
  if (!hmtx || hmtx.length < 4 * numberOfHMetrics || !cmapRaw) return null
  const cmap = parseCmap(cmapRaw)
  if (!cmap) return null
  const advances = new Uint16Array(numberOfHMetrics)
  for (let i = 0; i < numberOfHMetrics; i++) advances[i] = hmtx.readUInt16BE(4 * i)
  return { unitsPerEm, advances, cmapFormat: cmap.format, cmap: cmap.sub }
}

const styleTokensOf = (style: AdvanceStyle): string[] => [
  ...(style.bold ? ['bold'] : []),
  ...(style.italic ? ['italic'] : []),
]

/** parsed faces are a few hundred KB each — keep a handful per process */
const faceCache = new Map<string, ParsedFace | null>()
const FACE_CACHE_MAX = 8

function faceAdvances(face: FaceRef): ParsedFace | null {
  const key = `${face.path}#${face.offset}`
  const hit = faceCache.get(key)
  if (hit !== undefined) return hit
  let parsed: ParsedFace | null
  try {
    const fd = openSync(face.path, 'r')
    try {
      parsed = parseFaceAdvances(fd, face.offset)
    } finally {
      closeSync(fd)
    }
  } catch {
    return null // transient open/read failure: not cached, retried next call
  }
  if (faceCache.size >= FACE_CACHE_MAX) faceCache.clear()
  faceCache.set(key, parsed)
  return parsed
}

/**
 * Advance width of each codepoint of `text` rendered by the installed
 * family's face best matching `style`, in twips at `sizePt`. Unmapped
 * codepoints yield NaN (their real width depends on the renderer's fallback
 * font) — a NaN-poisoned sum tells callers the measurement is unusable.
 * Null when the family is not installed or carries no usable cmap/hmtx.
 */
export function advanceWidths(
  family: string,
  text: string,
  sizePt: number,
  style: AdvanceStyle = {},
): number[] | null {
  const candidates = getFontIndex().byFamily.get(norm(family))
  if (!candidates?.length) return null
  const want = styleTokensOf(style)
  const face = [...candidates].sort((a, b) => styleScore(b, want) - styleScore(a, want))[0]!
  const parsed = faceAdvances(face)
  if (!parsed) return null
  const scale = (sizePt * 20) / parsed.unitsPerEm
  const out: number[] = []
  for (const ch of text) {
    const cp = ch.codePointAt(0)!
    const glyph = glyphOf(parsed, cp)
    if (glyph === 0 && cp !== 0) {
      out.push(NaN)
      continue
    }
    const aw = parsed.advances[Math.min(glyph, parsed.advances.length - 1)]!
    out.push(aw * scale)
  }
  return out
}
