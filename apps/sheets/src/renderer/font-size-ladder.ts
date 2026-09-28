/** Excel's grow/shrink font walks its size ladder, not +-1. */
export const FONT_SIZE_LADDER = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 26, 28, 36, 48, 72]

export function stepFontSize(current: number, direction: 1 | -1): number {
  if (direction === 1) {
    return FONT_SIZE_LADDER.find((size) => size > current) ?? FONT_SIZE_LADDER.at(-1) ?? current
  }
  return (
    [...FONT_SIZE_LADDER].reverse().find((size) => size < current) ?? FONT_SIZE_LADDER[0] ?? current
  )
}
