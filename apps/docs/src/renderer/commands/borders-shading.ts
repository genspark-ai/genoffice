import { shdPatternDisplay, type ParaBorderLine } from '@genoffice/docx-engine'

export const BORDER_SIDES = ['t', 'b', 'l', 'r'] as const
export type BorderSide = (typeof BORDER_SIDES)[number]

/** the drawn sides of a paragraph, each with its declared look */
export type BorderSpec = Partial<Record<BorderSide, ParaBorderLine>>

export interface ShadingSpec {
  fill: string | null
  /** w:shd w:val; null is a plain fill */
  pattern: string | null
  /** the pattern's own colour; null is automatic */
  color: string | null
}

const HEX = /^[0-9A-Fa-f]{6}$/
const NAME = /^[A-Za-z0-9]+$/

function parseObject(json: string | null | undefined): Record<string, unknown> | null {
  if (!json) return null
  try {
    const parsed: unknown = JSON.parse(json)
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, unknown>) : null
  } catch {
    return null
  }
}

function lineOf(raw: unknown): ParaBorderLine {
  const src = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const line: ParaBorderLine = {}
  if (typeof src.color === 'string' && HEX.test(src.color)) line.color = src.color
  if (typeof src.szPt === 'number' && src.szPt > 0) line.szPt = src.szPt
  if (typeof src.spacePt === 'number' && src.spacePt > 0)
    line.spacePt = Math.min(31, Math.round(src.spacePt))
  if (typeof src.val === 'string' && NAME.test(src.val) && src.val !== 'single') line.val = src.val
  if (src.shadow === true) line.shadow = true
  if (src.frame === true) line.frame = true
  return line
}

/** what a paragraph's `borders` and `borderLines` attributes draw */
export function borderSpecOf(borders: unknown, borderLines: unknown): BorderSpec {
  const drawn = typeof borders === 'string' ? borders : ''
  const lines = parseObject(typeof borderLines === 'string' ? borderLines : null)
  const spec: BorderSpec = {}
  for (const side of BORDER_SIDES) if (drawn.includes(side)) spec[side] = lineOf(lines?.[side])
  return spec
}

/** the paragraph attributes that draw the spec */
export function borderAttrsOf(spec: BorderSpec): {
  borders: string | null
  borderLines: string | null
} {
  const lines: BorderSpec = {}
  let drawn = ''
  for (const side of BORDER_SIDES) {
    if (!spec[side]) continue
    drawn += side
    const line = lineOf(spec[side])
    if (Object.keys(line).length > 0) lines[side] = line
  }
  return {
    borders: drawn || null,
    borderLines: Object.keys(lines).length > 0 ? JSON.stringify(lines) : null,
  }
}

/** a border spec sent as a command payload; null when it is not one */
export function borderSpecOfPayload(payload: string): BorderSpec | null {
  const parsed = payload.startsWith('{') ? parseObject(payload) : null
  if (!parsed) return null
  const spec: BorderSpec = {}
  for (const side of BORDER_SIDES) if (parsed[side]) spec[side] = lineOf(parsed[side])
  return spec
}

export function shadingSpecOf(fill: unknown, shadingPattern: unknown): ShadingSpec {
  const pattern = parseObject(typeof shadingPattern === 'string' ? shadingPattern : null)
  const val = typeof pattern?.val === 'string' && NAME.test(pattern.val) ? pattern.val : null
  const color = typeof pattern?.color === 'string' && HEX.test(pattern.color) ? pattern.color : null
  return {
    fill: typeof fill === 'string' && HEX.test(fill) ? fill : null,
    pattern: val === 'clear' || val === 'nil' ? null : val,
    color: val ? color : null,
  }
}

/** the one colour the page shows for a shading: a pattern is its ink blended over the fill */
export function shadingDisplayOf(spec: ShadingSpec): string | null {
  if (!spec.pattern) return spec.fill
  return shdPatternDisplay(spec.pattern, spec.color ?? undefined, spec.fill ?? undefined) ?? null
}

export function shadingAttrsOf(spec: ShadingSpec): {
  shadingFill: string | null
  shadingPattern: string | null
  shadingDisplay: string | null
} {
  const display = shadingDisplayOf(spec)
  return {
    shadingFill: spec.fill,
    shadingPattern: spec.pattern
      ? JSON.stringify({ val: spec.pattern, ...(spec.color ? { color: spec.color } : {}) })
      : null,
    shadingDisplay: display && display !== spec.fill ? display : null,
  }
}

export function shadingSpecOfPayload(payload: string): ShadingSpec | null {
  const parsed = payload.startsWith('{') ? parseObject(payload) : null
  if (!parsed) return null
  return shadingSpecOf(
    parsed.fill,
    typeof parsed.pattern === 'string'
      ? JSON.stringify({ val: parsed.pattern, color: parsed.color })
      : null,
  )
}
