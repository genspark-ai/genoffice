/** ST_TextFontSize range (hundredths of a point, 100..400000) in points. */
export const FONT_SIZE_PT_MIN = 1
export const FONT_SIZE_PT_MAX = 4000

/** PowerPoint's grow/shrink ladder (pt); the ribbon size list shows the same values. */
export const FONT_SIZES = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96,
]

/** Runs whose size cannot be resolved step from PowerPoint's body default. */
export const DEFAULT_FONT_SIZE_PT = 18

/** Relative size change: 'ladder' = grow/shrink font (⇧⌘> / ⇧⌘<), 'point' = ±1 pt (⌘] / ⌘[). */
export interface FontSizeStep {
  dir: 1 | -1
  mode: 'ladder' | 'point'
}

const LADDER_MIN = FONT_SIZES[0]!
const LADDER_MAX = FONT_SIZES[FONT_SIZES.length - 1]!
// Beyond 96 pt PowerPoint keeps the ladder's last gap; below 8 pt it moves one point at a time
const ABOVE_LADDER_STEP = 8

/** Next/previous ladder value; a size between two rungs snaps to the neighbouring rung. */
export function stepFontSizePt(cur: number, dir: 1 | -1): number {
  if (dir > 0) {
    if (cur >= LADDER_MAX) return Math.min(FONT_SIZE_PT_MAX, cur + ABOVE_LADDER_STEP)
    if (cur < LADDER_MIN) return Math.min(LADDER_MIN, cur + 1)
    return FONT_SIZES.find((s) => s > cur)!
  }
  if (cur > LADDER_MAX) return Math.max(LADDER_MAX, cur - ABOVE_LADDER_STEP)
  if (cur <= LADDER_MIN) return Math.max(FONT_SIZE_PT_MIN, cur - 1)
  for (let i = FONT_SIZES.length - 1; i >= 0; i--) if (FONT_SIZES[i]! < cur) return FONT_SIZES[i]!
  return LADDER_MIN
}

export function nudgeFontSizePt(cur: number, dir: 1 | -1): number {
  return Math.min(FONT_SIZE_PT_MAX, Math.max(FONT_SIZE_PT_MIN, cur + dir))
}

export function applyFontSizeStep(cur: number, step: FontSizeStep): number {
  return step.mode === 'ladder' ? stepFontSizePt(cur, step.dir) : nudgeFontSizePt(cur, step.dir)
}
