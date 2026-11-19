/** Minimal sfnt cmap reader: can the face map every char of a string?
    Formats 4 and 12 only — which covers every face in EDIT_FONT_PATHS. */

const u16 = (b: Buffer, o: number) => b.readUInt16BE(o)
const u32 = (b: Buffer, o: number) => b.readUInt32BE(o)

function tableOffset(font: Buffer, tag: string): number {
  const numTables = u16(font, 4)
  for (let i = 0; i < numTables; i++) {
    const rec = 12 + i * 16
    if (font.toString('latin1', rec, rec + 4) === tag) return u32(font, rec + 8)
  }
  return -1
}

interface Subtable {
  offset: number
  format: number
}

function cmapSubtable(font: Buffer): Subtable | null {
  const cmap = tableOffset(font, 'cmap')
  if (cmap < 0) return null
  const n = u16(font, cmap + 2)
  let best: Subtable | null = null
  for (let i = 0; i < n; i++) {
    const rec = cmap + 4 + i * 8
    const platform = u16(font, rec)
    const encoding = u16(font, rec + 2)
    if (platform !== 0 && !(platform === 3 && (encoding === 1 || encoding === 10))) continue
    const offset = cmap + u32(font, rec + 4)
    const format = u16(font, offset)
    if (format === 12) return { offset, format }
    if (format === 4 && !best) best = { offset, format }
  }
  return best
}

/** Glyph id for a codepoint, 0 = unmapped (.notdef) */
function glyphId(font: Buffer, sub: Subtable, cp: number): number {
  const o = sub.offset
  if (sub.format === 12) {
    const groups = u32(font, o + 12)
    for (let i = 0; i < groups; i++) {
      const g = o + 16 + i * 12
      if (cp < u32(font, g)) return 0 // groups are sorted ascending
      if (cp <= u32(font, g + 4)) return u32(font, g + 8) + (cp - u32(font, g))
    }
    return 0
  }
  if (cp > 0xffff) return 0
  const segX2 = u16(font, o + 6)
  const ends = o + 14
  const starts = ends + segX2 + 2
  const deltas = starts + segX2
  const rangeOffs = deltas + segX2
  for (let s = 0; s < segX2; s += 2) {
    if (cp > u16(font, ends + s)) continue
    if (cp < u16(font, starts + s)) return 0
    const ro = u16(font, rangeOffs + s)
    if (ro === 0) return (cp + font.readInt16BE(deltas + s)) & 0xffff
    const gi = u16(font, rangeOffs + s + ro + (cp - u16(font, starts + s)) * 2)
    return gi === 0 ? 0 : (gi + font.readInt16BE(deltas + s)) & 0xffff
  }
  return 0
}

const hasGlyph = (font: Buffer, sub: Subtable, cp: number): boolean => glyphId(font, sub, cp) !== 0

/** True when the font maps every drawn char of text (unparseable font = false).
    Only line breaks are excluded: Unicode spaces (NBSP, U+3000, …) are drawn like any
    glyph and must be mapped — \s would silently exempt them. */
export function fontCoversText(font: Buffer, text: string): boolean {
  const chars = [...text.replace(/[\r\n]/g, '')]
  if (chars.length === 0) return true
  try {
    const sub = cmapSubtable(font)
    return sub !== null && chars.every((c) => hasGlyph(font, sub, c.codePointAt(0)!))
  } catch {
    return false
  }
}
