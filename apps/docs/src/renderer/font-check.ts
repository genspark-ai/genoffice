import type { ParsedDoc } from '@genoffice/docx-engine'
import { BUNDLED_FONTS, cssFontFamily, isBundledFont } from './line-metrics'

const GENERIC = new Set(['serif', 'sans-serif', 'monospace'])

/** every distinct font family the document declares */
export function collectDocFonts(parsed: ParsedDoc): string[] {
  const names = new Set<string>()
  const add = (font?: string) => {
    if (font) names.add(font)
  }
  for (const b of parsed.blocks) {
    for (const r of b.runs ?? []) add(r.font)
    for (const tb of b.textboxes ?? []) {
      for (const p of tb.paras) for (const r of p.runs) add(r.font)
    }
  }
  add(parsed.docDefaults?.asciiFont)
  add(parsed.docDefaults?.eastAsiaFont)
  for (const s of parsed.styles.values()) add(s.display?.font)
  return [...names]
}

// document.fonts.check is useless here: Chromium answers true for any family
// (system-fallback rendering counts). Detect availability by measuring — an
// unresolvable family inherits the exact metrics of the generic fallback.
const SAMPLE = '永体宋黑Wmg10'

function ctx2d(): CanvasRenderingContext2D | null {
  try {
    return document.createElement('canvas').getContext('2d')
  } catch {
    return null
  }
}

function fontAvailableIn(cx: CanvasRenderingContext2D, family: string): boolean {
  const q = `"${family.replace(/"/g, '')}"`
  const width = (stack: string): number => {
    cx.font = `16px ${stack}`
    return cx.measureText(SAMPLE).width
  }
  return (
    width(`${q}, monospace`) !== width('monospace') ||
    width(`${q}, sans-serif`) !== width('sans-serif')
  )
}

const symbolCoverCache = new Map<string, boolean>()
/** outside the F020-F0FF range every symbol-encoded legacy font maps, so this always draws .notdef */
const PUA_TOFU = '\uE000'

/**
 * Family installed AND covering every char of text — macOS ships a Symbol.ttf
 * without the F0xx cmap. An uncovered PUA char renders the stack's .notdef
 * (whose metrics differ from any other stack's), so the only reliable probe is
 * comparing against a known .notdef in the same stack.
 */
export function symbolFontCovers(family: string, text: string): boolean {
  const key = `${family}\u0000${text}`
  const hit = symbolCoverCache.get(key)
  if (hit !== undefined) return hit
  const cx = ctx2d()
  let covers = false
  if (cx) {
    cx.font = `32px "${family.replace(/"/g, '')}"`
    const metrics = (s: string): string => {
      const m = cx.measureText(s)
      return [
        m.width,
        m.actualBoundingBoxLeft,
        m.actualBoundingBoxRight,
        m.actualBoundingBoxAscent,
        m.actualBoundingBoxDescent,
      ].join()
    }
    const tofu = metrics(PUA_TOFU)
    covers = [...text].every((ch) => ch <= ' ' || metrics(ch) !== tofu)
  }
  symbolCoverCache.set(key, covers)
  return covers
}

export interface FontSubstitution {
  name: string
  /** first available family in the fallback chain, null = only bundled fallback left */
  substitute: string | null
}

/**
 * Declared fonts that don't resolve on this system, with the family that
 * actually renders instead (substitution used to be silent).
 */
export function checkMissingFonts(names: string[]): FontSubstitution[] {
  if (typeof document === 'undefined') return []
  const cx = ctx2d()
  if (!cx) return []
  const out: FontSubstitution[] = []
  for (const name of names) {
    // a declared bundled face resolves in canvas yet still renders the subset fallback — report it
    if (!isBundledFont(name) && fontAvailableIn(cx, name)) continue
    const chain = cssFontFamily(name)
      .split(',')
      .map((s) => s.trim().replace(/^'|'$/g, ''))
      .filter((c) => c !== name && !GENERIC.has(c))
    const substitute = chain.find((c) => !BUNDLED_FONTS.has(c) && fontAvailableIn(cx, c)) ?? null
    out.push({ name, substitute })
  }
  return out
}
