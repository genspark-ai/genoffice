import type {
  TextEffect,
  TextFill,
  TextFxOwn,
  TextGlow,
  TextReflection,
  TextShadow,
} from '@genoffice/docx-engine'
import { textHasCjk } from '../line-metrics'
import { darkPageBorderCss, darkPageColor, dkBackgroundImage } from './dark-page'

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

/** a gradient fill as a CSS image; OOXML turns clockwise from pointing right, CSS from pointing up */
export function textFillGradientCss(fill: Extract<TextFill, { kind: 'gradient' }>): string {
  const stops = fill.stops
    .map((stop) => `${cssColor(stop.color, stop.alpha)} ${Math.round(stop.pos * 1000) / 10}%`)
    .join(',')
  // one stop is one colour: CSS wants two to draw anything
  const list = fill.stops.length === 1 ? `${stops},${stops}` : stops
  return fill.path
    ? `radial-gradient(${fill.path === 'circle' ? 'circle' : 'ellipse'},${list})`
    : `linear-gradient(${Math.round((fill.angleDeg + 90) * 100) / 100}deg,${list})`
}

/**
 * w14:textFill on the glyphs. The colours are the document's; each gets a
 * `--dk-*` twin for the dark page. A gradient is clipped to the glyphs, so a
 * span that also paints a background (`onBackground`) keeps its plain colour.
 */
export function textFillDecls(fill: TextFill, onBackground = false): string[] {
  if (fill.kind === 'none') return ['-webkit-text-fill-color:transparent']
  if (fill.kind === 'solid') {
    const css = cssColor(fill.color, fill.alpha)
    return [`-webkit-text-fill-color:${css}`, `--dk-tf:${darkPageColor(css)}`]
  }
  if (onBackground) return []
  const image = textFillGradientCss(fill)
  return [
    `background-image:${image}`,
    '-webkit-background-clip:text',
    'background-clip:text',
    '-webkit-text-fill-color:transparent',
    dkBackgroundImage(image),
  ]
}

/** CSS declarations ("name:value") as a style object; custom properties keep their names */
export function declStyle(decls: string[]): Record<string, string> {
  const style: Record<string, string> = {}
  for (const decl of decls) {
    const at = decl.indexOf(':')
    const name = decl.slice(0, at)
    const key = name.startsWith('--')
      ? name
      : name.replace(/^-webkit-/, 'Webkit-').replace(/-([a-z])/g, (_, c: string) => c.toUpperCase())
    style[key] = decl.slice(at + 1)
  }
  return style
}

/** w14:textFill of a style: the rule's declarations and their dark-page twins */
export function styleTextFillDecls(
  fill: TextFill,
  opts: { outlined: boolean; character: boolean },
): { decls: string[]; darkDecls: string[] } {
  // glyphs with neither fill nor outline would vanish
  if (fill.kind === 'none')
    return { decls: opts.outlined ? ['-webkit-text-fill-color:transparent'] : [], darkDecls: [] }
  if (fill.kind === 'solid') {
    const css = cssColor(fill.color, fill.alpha)
    return {
      decls: [`-webkit-text-fill-color:${css}`],
      darkDecls: [`-webkit-text-fill-color:${darkPageColor(css)}`],
    }
  }
  // clipped to the glyphs it would take a paragraph's own shading with it
  if (!opts.character) return { decls: [], darkDecls: [] }
  const image = textFillGradientCss(fill)
  return {
    decls: [
      `background-image:${image}`,
      '-webkit-background-clip:text',
      'background-clip:text',
      '-webkit-text-fill-color:transparent',
    ],
    darkDecls: [`background-image:${darkPageBorderCss(image)}`],
  }
}

const validStop = (stop: unknown): boolean => {
  const s = stop as { pos?: unknown; color?: unknown } | null
  return !!s && typeof s.pos === 'number' && typeof s.color === 'string'
}

/** a TextFill as it was stored in an attribute; undefined when it is not one */
export function textFillOfJson(value: unknown): TextFill | undefined {
  const fill = value as Partial<TextFill> | null
  if (!fill || typeof fill !== 'object') return undefined
  if (fill.kind === 'none') return { kind: 'none' }
  if (fill.kind === 'solid' && typeof fill.color === 'string')
    return {
      kind: 'solid',
      color: fill.color,
      ...(typeof fill.alpha === 'number' ? { alpha: fill.alpha } : {}),
    }
  if (
    fill.kind === 'gradient' &&
    Array.isArray(fill.stops) &&
    fill.stops.length > 0 &&
    fill.stops.every(validStop)
  )
    return {
      kind: 'gradient',
      stops: fill.stops,
      angleDeg: typeof fill.angleDeg === 'number' ? fill.angleDeg : 0,
      ...(fill.path ? { path: fill.path } : {}),
    }
  return undefined
}

/**
 * The mark's `textFill` attr: the fill with the text colour it came with
 * (`over`, hex or null). A colour picked later replaces the fill, as in Word,
 * so the fill counts only while the colour is still that one.
 */
export function textFillAttr(fill: TextFill, color: string | null | undefined): string {
  return JSON.stringify({ ...fill, over: color ?? null })
}

export function liveTextFill(raw: unknown, color: unknown): TextFill | undefined {
  if (typeof raw !== 'string' || !raw) return undefined
  try {
    const stored = JSON.parse(raw) as { over?: string | null }
    const same = (stored.over ?? '').toUpperCase() === String(color ?? '').toUpperCase()
    return same ? textFillOfJson(stored) : undefined
  } catch {
    return undefined
  }
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

/** East Asian text breaks between any two glyphs; one box that long overflows the line or strands the run */
const SCALE_X_MAX_BREAKABLE_CHARS = 8

/** the inline-block cannot wrap: only a run short of break opportunities takes it */
export function charScaleXFits(text: string): boolean {
  if (/\s/.test(text)) return false
  return [...text].length <= SCALE_X_MAX_BREAKABLE_CHARS || !textHasCjk(text)
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
