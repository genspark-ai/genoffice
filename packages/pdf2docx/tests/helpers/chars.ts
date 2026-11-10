/** Hand-built PdfChar factories for analysis-layer unit tests (no wasm). */
import type { PdfChar } from '../../src/ir'
import { scriptOf } from '../../src/script'

export interface CharOpts {
  /** baseline y (PDF space, y up), default 700 */
  y?: number
  fontSize?: number
  /** glyph advance width; default 0.5 em (Latin) / 1 em (fullwidth) */
  width?: number
  fontFamily?: string
  fontWeight?: number
  italic?: boolean
  color?: string
  isGenerated?: boolean
  isHyphen?: boolean
  angle?: number
}

export function mkChar(text: string, x: number, opts: CharOpts = {}): PdfChar {
  const code = text.codePointAt(0) ?? 0
  const fontSize = opts.fontSize ?? 10
  const script = scriptOf(code)
  const defaultWidth =
    script === 'cjk' || script === 'kana' || script === 'hangul' ? fontSize : fontSize * 0.5
  const width = opts.width ?? defaultWidth
  const y = opts.y ?? 700
  return {
    code,
    text,
    box: { x0: x, x1: x + width, y0: y - fontSize * 0.21, y1: y + fontSize * 0.72 },
    looseBox: { x0: x, x1: x + width, y0: y - fontSize * 0.25, y1: y + fontSize * 0.95 },
    originX: x,
    originY: y,
    angle: opts.angle ?? 0,
    fontSize,
    fontWeight: opts.fontWeight ?? 400,
    fontFamily: opts.fontFamily ?? 'Helvetica',
    italic: opts.italic ?? false,
    color: opts.color ?? '000000',
    isGenerated: opts.isGenerated ?? false,
    isHyphen: opts.isHyphen ?? false,
    script,
  }
}

export interface TextOpts extends CharOpts {
  /** extra x gap between consecutive glyphs (tracking), default 0 */
  tracking?: number
  /** advance width of a space glyph in ems, default 0.25 */
  spaceEms?: number
}

/**
 * Lay out a string left-to-right starting at `x`. Spaces get their own char
 * (like a real space glyph in the PDF). Returns the chars and the x cursor
 * after the last glyph.
 */
export function mkText(
  text: string,
  x: number,
  opts: TextOpts = {},
): { chars: PdfChar[]; endX: number } {
  const fontSize = opts.fontSize ?? 10
  const tracking = opts.tracking ?? 0
  const chars: PdfChar[] = []
  let cursor = x
  for (const ch of text) {
    if (ch === ' ') {
      const w = (opts.spaceEms ?? 0.25) * fontSize
      chars.push(mkChar(' ', cursor, { ...opts, width: w }))
      cursor += w + tracking
      continue
    }
    const c = mkChar(ch, cursor, opts)
    chars.push(c)
    cursor += c.box.x1 - c.box.x0 + tracking
  }
  return { chars, endX: cursor }
}
