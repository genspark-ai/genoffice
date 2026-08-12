/**
 * Locate an installed font matching a PDF font's PostScript name (nameID 6) or family
 * (nameID 1/16), for text-edit rebuilds that keep the original typeface.
 */
import { readFileSync } from 'node:fs'

import { getFontIndex, norm, styleScore, styleTokens, OTTO_TAG, TTC_TAG } from './sfnt'

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
  if (!face) return null
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
