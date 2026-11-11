/** Calibrated measurement state and the label the file's Line annotation carries. */

export interface MeasureScale {
  /** how many PDF points one real-world unit spans */
  unitsPerPt: number
  unit: string
}

/** `235 pt` at 100 pt/m → `2.35 m`, always two decimals like measuring tools do */
export function formatLength(pts: number, scale: MeasureScale): string {
  const value = pts / scale.unitsPerPt
  return `${value.toFixed(2)} ${scale.unit}`
}

/** Validate a calibration entry: `value` is the real-world length the user
 *  says the drawn reference line (`drawnPts` PDF points) spans, so the line
 *  itself is the ruler — points per unit is its length over that entry. */
export function parseCalibration(
  value: string,
  unit: string,
  drawnPts: number,
): MeasureScale | null {
  const parsed = Number(value)
  if (!Number.isFinite(parsed) || parsed <= 0 || unit.trim() === '' || unit.length > 12) return null
  if (!Number.isFinite(drawnPts) || drawnPts <= 0) return null
  return { unitsPerPt: drawnPts / parsed, unit: unit.trim() }
}
