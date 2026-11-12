/**
 * Glyphless metric faces for the line strut. Chromium unions the strut's and
 * every inline box's half-leading geometry, so a Latin-led strut over East
 * Asian run spans pads each line past its line-height whenever the two faces
 * centre differently; Word sizes the line by the tallest face alone. The
 * PUA-blank source under the EA face's Word line split (hhea ascent + half the
 * remaining leading, see gridStrutMetricsPct) leads the strut and the spans
 * so every box shares one centre and the line is exactly the line-height.
 *
 * The document body face gets a stylesheet @font-face (doc-style-css); a
 * paragraph whose runs declare another East Asian face registers its own
 * FontFace here, one per distinct metric pair.
 */
import { PUA_BLANK_FAMILY } from './font-rules/face-classes'
import { gridStrutMetricsPct } from '../line-metrics'

export const GRID_STRUT_FAMILY = 'GenOffice Grid Strut'

/**
 * src of the bundled blank PUA face (fonts.css, build-hashed URL), read from the
 * loaded stylesheet so dev and packaged builds resolve the same file. Null (and
 * re-probed) until the sheet is present (tests/jsdom stay null).
 */
let blankSrcCache: string | null = null
export function blankFontFaceSrc(): string | null {
  if (blankSrcCache != null) return blankSrcCache
  try {
    for (const sheet of Array.from(document.styleSheets)) {
      let cssRules: CSSRuleList
      try {
        cssRules = sheet.cssRules
      } catch {
        continue // cross-origin sheet
      }
      for (const rule of Array.from(cssRules)) {
        const ff = rule as CSSFontFaceRule
        if (
          rule.type === CSSRule.FONT_FACE_RULE &&
          ff.style.getPropertyValue('font-family').includes(PUA_BLANK_FAMILY)
        ) {
          blankSrcCache = ff.style.getPropertyValue('src') || null
        }
      }
    }
  } catch {
    blankSrcCache = null
  }
  if (blankSrcCache == null) {
    // jsdom drops @font-face rules from the CSSOM: scan raw <style> text
    try {
      for (const el of Array.from(document.querySelectorAll('style'))) {
        const m = /@font-face\s*{[^}]*GenOffice PUA Blank[^}]*?src:\s*([^;}]+)/.exec(
          el.textContent ?? '',
        )
        if (m) blankSrcCache = m[1].trim()
      }
    } catch {
      blankSrcCache = null
    }
  }
  return blankSrcCache
}

const registered = new Map<string, string>()

/**
 * Family name of the registered strut face carrying `eaFace`'s line split, for
 * a paragraph's `--doc-strut-face`; null when the face has no resolver box or
 * the page has no FontFace API (the stylesheet strut, if any, then leads).
 */
export function declaredStrutFace(eaFace: string): string | null {
  const metrics = gridStrutMetricsPct(eaFace)
  if (!metrics) return null
  const key = `${metrics.ascentPct}/${metrics.descentPct}`
  const cached = registered.get(key)
  if (cached) return cached
  const src = blankFontFaceSrc()
  if (!src || typeof FontFace !== 'function' || typeof document === 'undefined' || !document.fonts)
    return null
  const family = `${GRID_STRUT_FAMILY} ${registered.size + 1}`
  try {
    const face = new FontFace(family, src, {
      ascentOverride: `${metrics.ascentPct}%`,
      descentOverride: `${metrics.descentPct}%`,
      lineGapOverride: '0%',
    })
    document.fonts.add(face)
    void face.load().catch(() => {})
  } catch {
    return null
  }
  registered.set(key, family)
  return family
}

/** test hook: forget the registered faces */
export function resetStrutFaces(): void {
  registered.clear()
  blankSrcCache = null
}
