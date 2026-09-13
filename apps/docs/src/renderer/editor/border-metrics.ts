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

/** closest CSS border-style; multi-line families become `double` at their total thickness */
export function borderCssStyle(style: string): string {
  if (style === 'dotted') return 'dotted'
  if (DASHED.has(style)) return 'dashed'
  return style in COMPOUND_PT ? 'double' : 'solid'
}
