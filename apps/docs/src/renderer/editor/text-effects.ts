import type {
  TextEffect,
  TextFxOwn,
  TextGlow,
  TextReflection,
  TextShadow,
} from '@genoffice/docx-engine'

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
  return `text-shadow:${glowShadows(glow)}`
}

function glowShadows(glow: TextGlow): string {
  const c = cssColor(glow.color, glow.alpha)
  return `0 0 ${glow.radiusPt / 2}pt ${c},0 0 ${glow.radiusPt}pt ${c}`
}

/** an outer shadow: thrown `distPt` along its direction (clockwise from pointing right), blurred */
function dropShadow(shadow: TextShadow): string {
  const rad = (shadow.dirDeg * Math.PI) / 180
  const round = (v: number) => Math.round(v * 100) / 100
  const x = round(Math.cos(rad) * shadow.distPt)
  const y = round(Math.sin(rad) * shadow.distPt)
  return `${x}pt ${y}pt ${shadow.blurPt}pt ${cssColor(shadow.color, shadow.alpha)}`
}

/** The glow and the shadow of a run are one CSS property: the shadow drawn over the halo. */
export function glowAndShadowDecl(glow?: TextGlow, shadow?: TextShadow): string | null {
  const layers = [shadow ? dropShadow(shadow) : '', glow ? glowShadows(glow) : ''].filter(Boolean)
  return layers.length > 0 ? `text-shadow:${layers.join(',')}` : null
}

export function parseTextShadowAttr(raw: string): TextShadow | undefined {
  try {
    const s = JSON.parse(raw) as Partial<TextShadow>
    if (
      typeof s.color !== 'string' ||
      typeof s.blurPt !== 'number' ||
      typeof s.distPt !== 'number' ||
      typeof s.dirDeg !== 'number'
    )
      return undefined
    return {
      color: s.color,
      ...(s.alpha !== undefined ? { alpha: s.alpha } : {}),
      blurPt: s.blurPt,
      distPt: s.distPt,
      dirDeg: s.dirDeg,
    }
  } catch {
    return undefined
  }
}

/**
 * The mirror image under the text. The mask is laid over the text's own box
 * before it is flipped, so its opaque end is the edge that touches the text.
 * The blur is not drawn.
 */
export function reflectionDecl(reflection: TextReflection): string {
  const clear = Math.round((1 - reflection.endPos) * 100)
  return (
    `-webkit-box-reflect:below ${reflection.distPt}pt ` +
    `linear-gradient(transparent ${clear}%,rgba(0,0,0,${reflection.startAlpha}))`
  )
}

export function parseReflectionAttr(raw: string): TextReflection | undefined {
  try {
    const r = JSON.parse(raw) as Partial<TextReflection>
    if (typeof r.startAlpha !== 'number' || typeof r.endPos !== 'number') return undefined
    return {
      startAlpha: r.startAlpha,
      endPos: r.endPos,
      blurPt: typeof r.blurPt === 'number' ? r.blurPt : 0,
      distPt: typeof r.distPt === 'number' ? r.distPt : 0,
    }
  } catch {
    return undefined
  }
}

/** the mark's `fxOwn` attr: the effects the editor set on the text itself */
export function parseFxOwnAttr(raw: unknown): TextFxOwn | undefined {
  if (typeof raw !== 'string' || !raw) return undefined
  try {
    const fx = JSON.parse(raw) as TextFxOwn
    return fx && typeof fx === 'object' ? fx : undefined
  } catch {
    return undefined
  }
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

/**
 * w:position moves the glyphs but leaves the line box alone: Word keeps the
 * line pitch for raised or lowered runs that share the line with an unshifted
 * run (probe 2026-11-03), so a relative offset rather than vertical-align.
 */
export function positionDecl(halfPoints: number): string {
  return `position:relative;top:${-halfPoints / 2}pt`
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

/** glyphs compress like Word; the negative margin hands the saved width back to the line.
 *  text-indent:0 — the block would inherit a hanging indent and shrink-to-fit
 *  to zero width, piling every glyph onto one spot */
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
    'text-indent:0',
  ]
}

/** text-align for a paragraph w:jc; distribute spreads every line, the last one included */
export function textAlignDecl(align: string): string {
  return align === 'distribute'
    ? 'text-align:justify;text-align-last:justify;text-justify:distribute'
    : `text-align:${align}`
}
