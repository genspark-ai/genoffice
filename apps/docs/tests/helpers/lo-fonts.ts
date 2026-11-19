/**
 * Real metrics from LO-bundled fonts (for the parity runner only).
 *
 * Baseline PDFs are laid out by LibreOffice, which ships metric-compatible
 * substitute fonts (Calibri→Carlito, Cambria→Caladea, Times/Arial/Courier→Liberation).
 * Parsing those font files for line-height/char-width metrics aligns us exactly
 * with the baseline's layout engine.
 *
 * Coverage: Latin fonts. System CJK fonts are all .ttc collections (unsupported
 * by opentype.js), and CJK char width is always 1.0em (em-square), so no real
 * files are needed; CJK line height still uses the simulateLines coefficient
 * path. Unmatched font families, or LO not installed, fall back to heuristics.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import * as opentype from 'opentype.js'
import {
  HeuristicMetrics,
  OpentypeMetrics,
  type FontMetricsProvider,
  type OpentypeFontLike,
  type RunStyle,
} from '../../src/renderer/line-metrics'

const LO_FONT_DIR = '/Applications/LibreOffice.app/Contents/Resources/fonts/truetype'
const WORD_DFONTS_DIR = '/Applications/Microsoft Word.app/Contents/Resources/DFonts'

/** Common Word Latin fonts → LO-bundled metric-compatible font file prefix */
const FAMILY_MAP: Array<[RegExp, string]> = [
  [/calibri|carlito/i, 'Carlito'],
  [/cambria|caladea/i, 'Caladea'],
  [/times|liberation ?serif/i, 'LiberationSerif'],
  [/^arial|helvetica|liberation ?sans/i, 'LiberationSans'],
  [/courier|liberation ?mono/i, 'LiberationMono'],
]

/** Word for Mac bundled fonts (DFonts): measure from the real file when the declared font is installed */
const DFONT_MAP: Array<[RegExp, { r: string; b: string; i: string; bi: string }]> = [
  [/garamond/i, { r: 'GARA.ttf', b: 'GARABD.ttf', i: 'GARAIT.ttf', bi: 'GARABDIT.ttf' }],
]

const cache = new Map<string, OpentypeFontLike | null>()

function loadFontFile(path: string): OpentypeFontLike | null {
  const hit = cache.get(path)
  if (hit !== undefined) return hit
  let wrapped: OpentypeFontLike | null = null
  if (existsSync(path)) {
    try {
      const buf = readFileSync(path)
      const font = opentype.parse(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength))
      const lineGap = (font.tables.hhea as { lineGap?: number } | undefined)?.lineGap ?? 0
      wrapped = {
        unitsPerEm: font.unitsPerEm,
        // Single line spacing = hhea ascent+descent+lineGap; fold lineGap into ascender
        ascender: font.ascender + lineGap,
        descender: font.descender,
        getAdvanceWidth: (t: string, s: number) => font.getAdvanceWidth(t, s),
        charToGlyphIndex: (ch: string) => font.charToGlyphIndex(ch),
      }
    } catch {
      wrapped = null
    }
  }
  cache.set(path, wrapped)
  return wrapped
}

function loadFont(prefix: string, bold: boolean, italic: boolean): OpentypeFontLike | null {
  const variant = bold && italic ? 'BoldItalic' : bold ? 'Bold' : italic ? 'Italic' : 'Regular'
  return loadFontFile(join(LO_FONT_DIR, `${prefix}-${variant}.ttf`))
}

function resolve(style: RunStyle): OpentypeFontLike | undefined {
  const prefix = FAMILY_MAP.find(([re]) => re.test(style.fontFamily))?.[1]
  if (prefix) return loadFont(prefix, style.bold, style.italic) ?? undefined
  const dfont = DFONT_MAP.find(([re]) => re.test(style.fontFamily))?.[1]
  if (dfont) {
    const file = style.bold && style.italic ? dfont.bi : style.bold ? dfont.b : style.italic ? dfont.i : dfont.r
    return loadFontFile(join(WORD_DFONTS_DIR, file)) ?? undefined
  }
  return undefined // Unknown font family (incl. CJK family names) → heuristics
}

/** Use real font metrics when LO is installed, else fall back entirely to heuristics (CI-safe) */
export function loBaselineMetrics(): FontMetricsProvider {
  if (!existsSync(LO_FONT_DIR)) return new HeuristicMetrics()
  return new OpentypeMetrics(resolve)
}
