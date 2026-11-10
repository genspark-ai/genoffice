/**
 * Word merges adjacent paragraphs whose border and shading settings are
 * identical into one border group (ECMA-376 §17.3.1.24): top/bottom lines draw
 * only at the group's edges (no w:between = no inner lines). Display-only —
 * the merge is applied as decoration classes; the borders attrs (which the
 * save path serializes back to w:pBdr) are never touched.
 */
import type { ParaBorderLine, ParaBorderSides } from '@genoffice/docx-engine'

/** the paragraph attrs Word's border-group merging compares */
export interface ParaBorderAttrs {
  borders?: string | null
  /** JSON per-side {color?,szPt?} as stored in the borderLines attr */
  borderLines?: string | null
  /** JSON per-side pt as stored in the borderPad attr (none sides that still pad) */
  borderPad?: string | null
  /** direct none/nil sides (cancel the style's side) */
  borderReset?: string | null
  shadingFill?: string | null
  /** pattern-shading display blend: pctNN paragraphs share a raw fill (auto)
   *  but differ visually, so the group comparison must see the blend */
  shadingDisplay?: string | null
  /** direct w:shd cancelling the style's fill */
  shadingClear?: boolean | null
  styleId?: string | null
  /** a non-paragraph block (table...): never part of a group, splits its neighbours */
  break?: true
}

/** the style-level side/fill a paragraph inherits (resolved through basedOn) */
export interface StyleBorderInfo {
  borderSides?: ParaBorderSides
  shadingFill?: string
}
/** null styleId = the w:default paragraph style (Normal), like Word's unstyled paragraphs */
export type StyleBorderLookup = (styleId: string | null) => StyleBorderInfo | undefined

const SIDES = ['t', 'b', 'l', 'r'] as const
type Side = (typeof SIDES)[number]

function parseJson<T>(raw: string | null | undefined): T | undefined {
  if (!raw) return undefined
  try {
    return JSON.parse(raw) as T
  } catch {
    return undefined
  }
}

function lineKey(line: ParaBorderLine | null | undefined): string {
  if (!line) return ''
  if (line.none) return line.spacePt ? `pad/${line.spacePt}` : ''
  return `${line.color ?? ''}/${line.szPt ?? ''}/${line.spacePt ?? ''}`
}

/**
 * Canonical per-side border + fill of a paragraph, direct pPr winning per side
 * over the style (Word merges w:pBdr side by side). '' = nothing to group on.
 */
export function borderGroupKey(a: ParaBorderAttrs, lookup?: StyleBorderLookup): string {
  if (a.break) return ''
  const style = lookup?.(a.styleId ?? null)
  const lines = parseJson<Partial<Record<Side, ParaBorderLine>>>(a.borderLines)
  const pad = parseJson<Partial<Record<Side, number>>>(a.borderPad)
  const drawn = typeof a.borders === 'string' ? a.borders : ''
  const reset = typeof a.borderReset === 'string' ? a.borderReset : ''
  let sides = ''
  for (const side of SIDES) {
    let key: string
    if (drawn.includes(side)) key = lineKey(lines?.[side] ?? {})
    else if (pad?.[side]) key = `pad/${pad[side]}`
    else if (reset.includes(side)) key = ''
    else key = lineKey(style?.borderSides?.[side])
    sides += `${side}:${key}|`
  }
  if (!sides.replace(/[tblr]:\|/g, '')) return ''
  const fill =
    a.shadingDisplay ??
    a.shadingFill ??
    (a.shadingClear
      ? null
      : style?.shadingFill && style.shadingFill !== 'auto'
        ? style.shadingFill
        : null)
  return `${sides}fill:${fill ?? ''}`
}

/** a touched paragraph whose attrs can form or change a border group */
export function affectsBorderGroup(a: ParaBorderAttrs, lookup?: StyleBorderLookup): boolean {
  return !!(
    a.borders ||
    a.borderLines ||
    a.borderPad ||
    a.borderReset ||
    a.shadingFill ||
    a.shadingDisplay ||
    a.shadingClear ||
    borderGroupKey(a, lookup)
  )
}

export function sameBorderGroup(
  a: ParaBorderAttrs,
  b: ParaBorderAttrs,
  lookup?: StyleBorderLookup,
): boolean {
  const ka = borderGroupKey(a, lookup)
  return ka !== '' && ka === borderGroupKey(b, lookup)
}

/**
 * Per-paragraph border suppression for a run of adjacent top-level paragraphs:
 * same group as the next → its bottom border is an inner boundary (suppress);
 * same group as the previous → suppress its top border.
 */
export function borderMergeFlags(
  paras: ParaBorderAttrs[],
  lookup?: StyleBorderLookup,
): Array<{ suppressTop: boolean; suppressBottom: boolean }> {
  const keys = paras.map((p) => borderGroupKey(p, lookup))
  const same = (i: number, j: number) => keys[i] !== '' && keys[i] === keys[j]
  return paras.map((_, i) => ({
    suppressTop: i > 0 && same(i - 1, i),
    suppressBottom: i < paras.length - 1 && same(i, i + 1),
  }))
}
