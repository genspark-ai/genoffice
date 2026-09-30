/**
 * Word's drawn thickness of an OOXML border line. w:sz (eighths of a point)
 * is the width of ONE line; compound styles add fixed 0.75pt "thin" strokes
 * and gaps around it (probed against Word 2026-09-12: single/double/triple
 * and the thin/thick families at sz 4, 12, 24, 48). Rows advance by this total,
 * so every layout consumer must measure the compound thickness, not w:sz.
 */
export type BorderLine = { style: string; szEighths?: number } | undefined | null

const THIN_PT = 0.75

const COMPOUND_PT: Record<string, (w: number) => number> = {
  double: (w) => 3 * w,
  triple: (w) => 5 * w,
  thinThickSmallGap: (w) => w + 2 * THIN_PT,
  thickThinSmallGap: (w) => w + 2 * THIN_PT,
  thinThickThinSmallGap: (w) => w + 4 * THIN_PT,
  thinThickMediumGap: (w) => 2 * w,
  thickThinMediumGap: (w) => 2 * w,
  thinThickThinMediumGap: (w) => 3 * w,
  thinThickLargeGap: (w) => w + 3 * THIN_PT,
  thickThinLargeGap: (w) => w + 3 * THIN_PT,
  thinThickThinLargeGap: (w) => 2 * w + 4 * THIN_PT,
}

const DASHED = new Set(['dashed', 'dashSmallGap', 'dotDash', 'dotDotDash', 'dashDotStroked'])

export function isDrawnBorder(b: BorderLine): b is { style: string; szEighths?: number } {
  return !!b && b.style !== 'none' && b.style !== 'nil'
}

export function borderTotalPt(b: BorderLine): number {
  if (!isDrawnBorder(b)) return 0
  const w = (b.szEighths ?? 4) / 8
  return COMPOUND_PT[b.style]?.(w) ?? w
}

/** unrounded CSS px of the drawn thickness */
export function borderTruePx(b: BorderLine): number {
  return (borderTotalPt(b) / 72) * 96
}

/** whole px we draw (Chromium snaps sub-px borders anyway), at least 1 for a drawn line */
export function borderDrawnPx(b: BorderLine): number {
  return isDrawnBorder(b) ? Math.max(1, Math.round(borderTruePx(b))) : 0
}

/** drawn px of the collapsed line between a cell and its neighbour: the neighbour's line
 *  still occupies half of this cell's box when the cell's own side is nil; undefined =
 *  no per-cell evidence, the table-level default applies */
export function collapsedEdgePx(own: BorderLine, neighbour: BorderLine): number | undefined {
  if (isDrawnBorder(own)) return borderDrawnPx(own)
  if (isDrawnBorder(neighbour)) return borderDrawnPx(neighbour)
  return own ? 0 : undefined
}

/** cell margin twips as a CSS px length without a rounded-off fraction (0.667px of a 0.5pt border matters) */
export function cellPadPx(twips: number): string {
  return `${Math.round((twips / 15) * 1000) / 1000}px`
}

/** closest CSS border-style; multi-line families become `double` at their total thickness */
export function borderCssStyle(style: string): string {
  if (style === 'dotted') return 'dotted'
  if (DASHED.has(style)) return 'dashed'
  return style in COMPOUND_PT ? 'double' : 'solid'
}

/** One cell diagonal as a CSS background layer, for w:tl2br / w:tr2bl.
 *
 *  CSS "magic corner" keywords aim the gradient at a corner, so the hard-stop
 *  band runs along that corner's diagonal at ANY cell aspect ratio (a plain
 *  `45deg` gradient would not): `to bottom right` bands along top-left to
 *  bottom-right, `to top right` along the other one. Band stops are px offsets
 *  along the gradient axis, whose perpendicular distance is the drawn width.
 *
 *  Only the weight and color carry over; a gradient cannot express the dashed
 *  and compound families, so those draw solid, as Word's own preview does. */
export function cellDiagonalCss(
  b: { style: string; szEighths?: number; color?: string } | undefined | null,
  which: 'tl2br' | 'tr2bl',
): string | null {
  if (!isDrawnBorder(b)) return null
  const half = borderDrawnPx(b) / 2
  const color = b.color && b.color !== 'auto' ? `#${b.color}` : '#000'
  const dir = which === 'tl2br' ? 'to bottom right' : 'to top right'
  const stops = [
    `transparent calc(50% - ${half}px)`,
    `${color} calc(50% - ${half}px)`,
    `${color} calc(50% + ${half}px)`,
    `transparent calc(50% + ${half}px)`,
  ]
  return `linear-gradient(${dir},${stops.join(',')})`
}
