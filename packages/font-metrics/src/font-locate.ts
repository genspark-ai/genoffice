/**
 * Locate an installed font matching a PDF font's PostScript name (nameID 6) or family
 * (nameID 1/16), for text-edit rebuilds that keep the original typeface.
 */
import { closeSync, openSync, readFileSync } from 'node:fs'

import {
  getFontIndex,
  norm,
  readTable,
  readTableDir,
  styleScore,
  styleTokens,
  OTTO_TAG,
  TTC_TAG,
} from './sfnt'
import type { FaceRef } from './sfnt'

/** Picks FPDFText_LoadFont's font_type: glyf faces embed as FontFile2 (TRUETYPE),
    CFF faces as FontFile3 (TYPE1) */
export const isTruetype = (b: Buffer): boolean => b.length >= 4 && b.readUInt32BE(0) !== OTTO_TAG

/** Extract a single face from a ttc into a standalone sfnt (rewrite the table directory,
    copy table data; passthrough for plain ttf) */
function extractFace(buf: Buffer, offset: number): Buffer {
  if (buf.readUInt32BE(0) !== TTC_TAG) return buf
  const numTables = buf.readUInt16BE(offset + 4)
  let total = 12 + 16 * numTables
  const entries: Array<{ dirPos: number; tOff: number; tLen: number; newOff: number }> = []
  for (let t = 0; t < numTables; t++) {
    const e = offset + 12 + 16 * t
    const tLen = buf.readUInt32BE(e + 12)
    entries.push({ dirPos: e, tOff: buf.readUInt32BE(e + 8), tLen, newOff: total })
    total += (tLen + 3) & ~3
  }
  const out = Buffer.alloc(total)
  buf.copy(out, 0, offset, offset + 12)
  for (let t = 0; t < numTables; t++) {
    const e = entries[t]!
    buf.copy(out, 12 + 16 * t, e.dirPos, e.dirPos + 8)
    out.writeUInt32BE(e.newOff, 12 + 16 * t + 8)
    out.writeUInt32BE(e.tLen, 12 + 16 * t + 12)
    buf.copy(out, e.newOff, e.tOff, e.tOff + e.tLen)
  }
  return out
}

/** The family (nameID 1/16, normalized-exact match) has at least one installed face. */
export function isFamilyInstalled(family: string): boolean {
  const key = norm(family)
  return key !== '' && getFontIndex().byFamily.has(key)
}

const faceCache = new Map<string, Buffer>()

/**
 * Standalone sfnt bytes of the installed font best matching (psName, family):
 * exact PostScript name first, then family with the PS name's style tokens
 * (Regular preferred when there are none). Null when nothing matches.
 */
export function findSystemFont(psName: string, family: string): Buffer | null {
  const index = getFontIndex()
  let face = psName ? index.byPs.get(norm(psName)) : undefined
  if (!face && family) {
    const candidates = index.byFamily.get(norm(family))
    if (candidates?.length) {
      const want = styleTokens(psName)
      face = [...candidates].sort((a, b) => styleScore(b, want) - styleScore(a, want))[0]
    }
  }
  return face ? faceBytes(face) : null
}

/** Standalone sfnt bytes of a face (cached; extracts from .ttc); null when unreadable */
function faceBytes(face: FaceRef): Buffer | null {
  const key = `${face.path}#${face.offset}`
  let bytes = faceCache.get(key)
  if (!bytes) {
    try {
      bytes = extractFace(readFileSync(face.path), face.offset)
    } catch {
      return null
    }
    if (faceCache.size >= 4) faceCache.clear()
    faceCache.set(key, bytes)
  }
  return bytes
}

// ── Full-index coverage scan ────────────────────────────────────────
// Last-resort fallback mirroring what the browser preview does: it resolves per-char
// fallback across EVERY installed font, so text that displays must also find an
// embeddable face here — or saving would fail for text the user already sees.

/** Best subtable of a standalone cmap table (offsets relative to the table start):
    format 12 preferred, else format 4 — same policy as the pdf app's cmap reader */
function cmapSubtable(cmap: Buffer): { off: number; format: number } | null {
  const n = cmap.readUInt16BE(2)
  let best: { off: number; format: number } | null = null
  for (let i = 0; i < n; i++) {
    const rec = 4 + i * 8
    if (rec + 8 > cmap.length) break
    const platform = cmap.readUInt16BE(rec)
    const encoding = cmap.readUInt16BE(rec + 2)
    if (platform !== 0 && !(platform === 3 && (encoding === 1 || encoding === 10))) continue
    const off = cmap.readUInt32BE(rec + 4)
    if (off + 4 > cmap.length) continue
    const format = cmap.readUInt16BE(off)
    if (format === 12) return { off, format }
    if (format === 4 && !best) best = { off, format }
  }
  return best
}

function mapsCodepoint(cmap: Buffer, sub: { off: number; format: number }, cp: number): boolean {
  const o = sub.off
  if (sub.format === 12) {
    const groups = cmap.readUInt32BE(o + 12)
    for (let i = 0; i < groups; i++) {
      const g = o + 16 + i * 12
      if (g + 12 > cmap.length) return false
      if (cp < cmap.readUInt32BE(g)) return false // groups sorted ascending
      if (cp <= cmap.readUInt32BE(g + 4))
        return cmap.readUInt32BE(g + 8) + (cp - cmap.readUInt32BE(g)) !== 0
    }
    return false
  }
  if (cp > 0xffff) return false
  const segX2 = cmap.readUInt16BE(o + 6)
  const ends = o + 14
  const starts = ends + segX2 + 2
  const deltas = starts + segX2
  const rangeOffs = deltas + segX2
  for (let s = 0; s < segX2; s += 2) {
    if (cp > cmap.readUInt16BE(ends + s)) continue
    if (cp < cmap.readUInt16BE(starts + s)) return false
    const ro = cmap.readUInt16BE(rangeOffs + s)
    if (ro === 0) return ((cp + cmap.readInt16BE(deltas + s)) & 0xffff) !== 0
    const gi = cmap.readUInt16BE(rangeOffs + s + ro + (cp - cmap.readUInt16BE(starts + s)) * 2)
    return gi !== 0 && ((gi + cmap.readInt16BE(deltas + s)) & 0xffff) !== 0
  }
  return false
}

/** Color-emoji faces display in the browser but cannot embed as PDF text objects */
const COLOR_FONT_TABLES = ['sbix', 'COLR', 'CBDT', 'CBLC']

/**
 * First installed face whose cmap maps every drawn char of `text`, plainest style
 * first (regular faces outrank bold/italic/decorative ones). Coverage is checked
 * from targeted cmap-table reads — CJK collections run 100MB+, so candidate faces
 * are never read whole; only the one winning face is extracted (and cached).
 * Null when nothing covers the text (e.g. emoji, which only color fonts map).
 */
export function findFontCovering(text: string): Buffer | null {
  const cps = [...new Set([...text.replace(/[\r\n]/g, '')].map((c) => c.codePointAt(0)!))]
  if (cps.length === 0) return null
  const index = getFontIndex()
  const seen = new Set<string>()
  const faces: FaceRef[] = []
  for (const face of [...index.byPs.values(), ...[...index.byFamily.values()].flat()]) {
    const key = `${face.path}#${face.offset}`
    if (seen.has(key)) continue
    seen.add(key)
    faces.push(face)
  }
  faces.sort((a, b) => styleScore(b, []) - styleScore(a, []))
  for (const face of faces) {
    let fd: number
    try {
      fd = openSync(face.path, 'r')
    } catch {
      continue
    }
    let covers = false
    try {
      const tables = readTableDir(fd, face.offset)
      if (
        tables &&
        (tables.has('glyf') || tables.has('CFF ')) &&
        !COLOR_FONT_TABLES.some((t) => tables.has(t))
      ) {
        const cmap = readTable(fd, tables, 'cmap', 1 << 22)
        const sub = cmap ? cmapSubtable(cmap) : null
        covers = sub !== null && cps.every((cp) => mapsCodepoint(cmap!, sub, cp))
      }
    } catch {
      covers = false
    } finally {
      closeSync(fd)
    }
    if (!covers) continue
    const bytes = faceBytes(face)
    if (bytes) return bytes
  }
  return null
}
