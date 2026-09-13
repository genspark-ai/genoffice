import type { TextEffect, TextGlow } from '@genoffice/docx-engine'

const TEXT_EFFECTS: ReadonlySet<string> = new Set(['outline', 'emboss', 'imprint', 'shadow'])

export function parseTextEffectAttr(raw: unknown): TextEffect | undefined {
  return typeof raw === 'string' && TEXT_EFFECTS.has(raw) ? (raw as TextEffect) : undefined
}

/** Word draws embossed/engraved glyphs in the paper colour: only the offset shadow shows */
export function paperColorEffect(effect: TextEffect | undefined): boolean {
  return effect === 'emboss' || effect === 'imprint'
}

export function textEffectDecls(effect: TextEffect): string[] {
  switch (effect) {
    case 'outline':
      return ['-webkit-text-stroke:0.5px currentColor', '-webkit-text-fill-color:transparent']
    case 'emboss':
      return ['color:var(--docs-paper)', 'text-shadow:1px 1px 0 var(--docs-paper-ink-mid)']
    case 'imprint':
      return ['color:var(--docs-paper)', 'text-shadow:-1px -1px 0 var(--docs-paper-ink-mid)']
    case 'shadow':
      return ['text-shadow:1px 1px 0 var(--docs-paper-ink-soft)']
  }
}

function cssColor(hex: string, alpha: number | undefined): string {
  const n = parseInt(hex, 16)
  return alpha !== undefined && alpha < 1
    ? `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${alpha})`
    : `#${hex}`
}

/** two blur layers: one Gaussian alone fades far below Word's dense halo */
export function glowDecl(glow: TextGlow): string {
  const c = cssColor(glow.color, glow.alpha)
  return `text-shadow:0 0 ${glow.radiusPt / 2}pt ${c},0 0 ${glow.radiusPt}pt ${c}`
}

export function doubleStrikeDecl(): string {
  return 'text-decoration:line-through double'
}

export function parseGlowAttr(raw: string): TextGlow | undefined {
  try {
    const g = JSON.parse(raw) as Partial<TextGlow>
    if (typeof g.color !== 'string' || typeof g.radiusPt !== 'number') return undefined
    return {
      color: g.color,
      radiusPt: g.radiusPt,
      ...(g.alpha !== undefined ? { alpha: g.alpha } : {}),
    }
  } catch {
    return undefined
  }
}

export function positionDecl(halfPoints: number): string {
  return `vertical-align:${halfPoints / 2}pt`
}

/** w:w on a whitespace-free run: {s: scale factor, gapEm: estimated width change in em} */
export interface CharScaleX {
  s: number
  gapEm: number
}

export function charScaleXAttr(text: string, scalePct: number, perCharEm: number): string {
  const chars = [...text].length
  return JSON.stringify({
    s: scalePct / 100,
    gapEm: Math.round(perCharEm * chars * 1000) / 1000,
  })
}

/** glyphs compress like Word; the negative margin hands the saved width back to the line */
export function charScaleXDecls(json: string): string[] {
  let v: Partial<CharScaleX>
  try {
    v = JSON.parse(json) as Partial<CharScaleX>
  } catch {
    return []
  }
  if (typeof v.s !== 'number' || typeof v.gapEm !== 'number' || !(v.s > 0)) return []
  return [
    'display:inline-block',
    `transform:scaleX(${v.s})`,
    'transform-origin:0 50%',
    `margin-right:${v.gapEm}em`,
  ]
}

/** text-align for a paragraph w:jc; distribute spreads every line, the last one included */
export function textAlignDecl(align: string): string {
  return align === 'distribute'
    ? 'text-align:justify;text-align-last:justify;text-justify:distribute'
    : `text-align:${align}`
}
