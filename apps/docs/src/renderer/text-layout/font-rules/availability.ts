/**
 * D1 of the font-rules inventory: is the declared face the one Chromium would
 * draw. Bundled stand-ins count as missing (Word lays the real family out, the
 * renderer draws a metric twin, so every fontTable hint still applies).
 */
import { isEastAsianFaceName as isCjkFontName } from '../face-info'

/**
 * Font family → canvas CSS font-family fallback chain.
 * When a common Word font is missing locally, fall back to metric-compatible
 * open-source fonts (registered in fonts/fonts.css); identical widths keep line
 * breaks aligned with Word and the offline pagination model. CJK families fall
 * back to macOS equivalents (CJK width is always 1em, so this is mostly glyph appearance).
 */
/**
 * Whether a font family actually resolves on this machine. document.fonts.check
 * is useless in Chromium (true for any unknown system-ish family), so this
 * measures a mixed-script sample against both generic fallbacks: a family that
 * changes neither width doesn't exist.
 */
/**
 * Bundled web fonts (fonts.css). The canvas probe resolves them, but they are
 * last-resort subset faces, so substitution decisions treat them as missing.
 */
export const BUNDLED_FONTS = new Set([
  'Noto Sans CJK SC',
  'Noto Serif CJK SC',
  'GenOffice Sans KR',
  'GenOffice Serif KR',
  'GenOffice Gothic KR',
  'GenOffice Poppins',
  'GenOffice DM Sans',
  'GenOffice Tamil',
  'GenOffice Cordia',
  'GenOffice Sarabun',
  'GenOffice Bengali',
  'GenOffice Fullwidth TC',
  'GenOffice Songti SC',
  'GenOffice Songti TC',
  'Carlito GO',
  'Aptos GO',
  'Aptos Display GO',
  'Aptos Narrow GO',
  'Century Gothic GO',
  'Cambria Math GO',
  'Caladea',
  'Liberation Serif',
  'Liberation Sans',
  'Liberation Mono',
])
const BUNDLED_LC = new Set([...BUNDLED_FONTS].map((f) => f.toLowerCase()))
export function isBundledFont(name: string): boolean {
  return BUNDLED_LC.has(name.toLowerCase())
}

/**
 * The declared face is not the one Chromium would draw: absent from this
 * machine, or one of the bundled stand-ins (Word lays the real family out,
 * the renderer draws a metric twin, so the fontTable hints still apply).
 * Every "missing face" rule gates on this one test.
 */
export function declaredFaceMissing(head: string): boolean {
  return isBundledFont(head) || !isFontAvailable(head)
}
/** a CJK-named face this machine draws with a stand-in (the grid strut's reason to exist) */
export function missingCjkFace(head: string): boolean {
  return isCjkFontName(head) && declaredFaceMissing(head)
}

const fontAvailableCache = new Map<string, boolean>()

/** embedded faces registered/revoked: their availability may have flipped */
export function forgetFontAvailability(families: readonly string[]): void {
  for (const f of families) fontAvailableCache.delete(f)
}
export function isFontAvailable(font: string): boolean {
  if (typeof document === 'undefined') return false
  const cached = fontAvailableCache.get(font)
  if (cached !== undefined) return cached
  let available = false
  try {
    const canvas = document.createElement('canvas')
    const ctx = canvas.getContext('2d')
    if (ctx) {
      const sample = '한글あア中文abcWXYmm123'
      const quoted = `"${font.replace(/"/g, '')}"`
      const widthWith = (family: string) => {
        ctx.font = `32px ${family}`
        return ctx.measureText(sample).width
      }
      available =
        widthWith(`${quoted}, monospace`) !== widthWith('monospace') ||
        widthWith(`${quoted}, serif`) !== widthWith('serif')
    }
  } catch {
    /* headless/test environments treat every font as missing */
  }
  fontAvailableCache.set(font, available)
  return available
}

export const GENERIC_FAMILIES = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
])
