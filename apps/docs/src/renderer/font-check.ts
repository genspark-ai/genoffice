import type { ParsedDoc } from '@genoffice/docx-engine'
import { cssFontFamily } from './line-metrics'

/** bundled web fonts (fonts.css) — always "available" but they are the
 *  one-size-fits-all fallback, so landing on them still counts as a miss */
const BUNDLED = new Set([
  'Noto Sans CJK SC',
  'Noto Serif CJK SC',
  'Carlito',
  'Caladea',
  'Liberation Serif',
  'Liberation Sans',
  'Liberation Mono',
])

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
    if (fontAvailableIn(cx, name)) continue
    const chain = cssFontFamily(name)
      .split(',')
      .map((s) => s.trim().replace(/^'|'$/g, ''))
      .filter((c) => c !== name && !GENERIC.has(c))
    const substitute = chain.find((c) => !BUNDLED.has(c) && fontAvailableIn(cx, c)) ?? null
    out.push({ name, substitute })
  }
  return out
}
