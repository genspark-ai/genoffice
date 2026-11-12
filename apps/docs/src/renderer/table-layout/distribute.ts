import { MIN_COL_TWIPS } from './constants'

/**
 * Word autofit for a table whose cells carry no w:tcW (probe 2026-10-20, pct
 * 100/50 tables, uniform and skewed grids): the grid is ignored and the table
 * width is shared like CSS auto layout - over the max-content sum the extra is
 * split in proportion to each column's max-content; under it every column
 * starts at its min-content and the rest is split by (max - min).
 */
export function distributeContentAutofit(
  tableTwips: number,
  mins: number[],
  maxes: number[],
): number[] {
  const n = mins.length
  const lo = mins.map((m) => Math.max(1, m))
  const hi = maxes.map((m, i) => Math.max(m, lo[i]))
  const sumLo = lo.reduce((a, b) => a + b, 0)
  const sumHi = hi.reduce((a, b) => a + b, 0)
  let widths: number[]
  if (tableTwips >= sumHi) {
    const extra = tableTwips - sumHi
    widths = hi.map((h) => h + (sumHi > 0 ? (extra * h) / sumHi : extra / n))
  } else if (tableTwips > sumLo) {
    const room = tableTwips - sumLo
    const slack = hi.map((h, i) => h - lo[i])
    const sumSlack = slack.reduce((a, b) => a + b, 0)
    widths = lo.map((l, i) => l + (sumSlack > 0 ? (room * slack[i]) / sumSlack : room / n))
  } else {
    const scale = sumLo > 0 ? tableTwips / sumLo : 1
    widths = lo.map((l) => l * scale)
  }
  const rounded = widths.map((w) => Math.max(1, Math.round(w)))
  const drift = tableTwips - rounded.reduce((a, b) => a + b, 0)
  if (drift !== 0 && rounded.length) {
    let widest = 0
    rounded.forEach((w, i) => {
      if (w > rounded[widest]) widest = i
    })
    rounded[widest] = Math.max(1, rounded[widest] + drift)
  }
  return rounded
}

/** the autofit reclaim pass never compresses a column below its min-content
 * or MIN_COL_TWIPS; sliver slots already under that keep their width */
export function reclaimFloorTwips(width: number, minContent: number): number {
  return Math.max(minContent, Math.min(width, MIN_COL_TWIPS))
}

/** a nested table over its cell width shrinks each column to its min-content;
 * an empty column keeps its side margins (Word: tcW 360/360 holding a nested
 * table and "8" in a 544-twip cell -> 222/328), sliver grid slots keep their width */
export function nestedFloorTwips(width: number, minContent: number, sideMar: number): number {
  return Math.max(minContent, Math.min(width, sideMar))
}

/** take `overflow` twips out of the columns in proportion to their surplus over `floors` */
export function reclaimSurplus(widths: number[], floors: number[], overflow: number): number[] {
  const surplus = widths.map((w, i) => Math.max(0, w - floors[i]))
  const totalSurplus = surplus.reduce((a, b) => a + b, 0)
  if (totalSurplus <= 0) return widths
  const take = Math.min(overflow, totalSurplus)
  return widths.map((w, i) => w - Math.round((surplus[i] / totalSurplus) * take))
}

/** min-contents alone can exceed the budget (one huge unbreakable token). Word
 * keeps every column's side margins and scales only the text widths (probe
 * 2026-11-09: 40 m / 5 m / 5 m / 5 m on a 9026 column -> 6144 / 960 / 955 / 955,
 * = 216 + text x 0.845; tcW and max-content never enter); `floors` are those
 * margins. Without floors, or when they alone overflow, the whole widths scale. */
export function scaleToFit(widths: number[], budget: number, floors?: number[]): number[] {
  const sum = widths.reduce((a, b) => a + b, 0)
  if (sum <= budget) return widths
  const base = floors?.map((f, i) => Math.min(Math.max(f, 0), widths[i])) ?? widths.map(() => 0)
  const sumBase = base.reduce((a, b) => a + b, 0)
  const scaled =
    sumBase >= budget || sumBase >= sum
      ? widths.map((w) => (w * budget) / sum)
      : widths.map((w, i) => base[i] + ((w - base[i]) * (budget - sumBase)) / (sum - sumBase))
  const rounded = scaled.map((w) => Math.max(1, Math.round(w)))
  // rounding drift goes to the widest column so the table keeps its width
  const drift = budget - rounded.reduce((a, b) => a + b, 0)
  if (drift !== 0) {
    const widest = rounded.indexOf(Math.max(...rounded))
    rounded[widest] = Math.max(1, rounded[widest] + drift)
  }
  return rounded
}
