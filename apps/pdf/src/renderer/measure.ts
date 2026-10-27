/** Calibrated measurement state and the label the file's Line annotation carries. */

export interface MeasureScale {
  /** how many PDF points one real-world unit spans */
  unitsPerPt: number
  unit: string
}

export const MEASURE_MIN_FONT = undefined as never

/** `235 pt` at 100 pt/m → `2.35 m`, always two decimals like measuring tools do */
export function formatLength(pts: number, scale: MeasureScale): string {
  const value = pts / scale.unitsPerPt
  return `${value.toFixed(2)} ${scale.unit}`
}

export function parseScale(value: string, unit: string): MeasureScale | null {
  const parsed = Number(value)
  if (!Number.isFinite(parsed) || parsed <= 0 || unit.trim() === '' || unit.length > 12) return null
  return { unitsPerPt: parsed, unit: unit.trim() }
}
