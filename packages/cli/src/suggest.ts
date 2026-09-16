import type { ErrorHints } from './result'

/** Optimal-string-alignment distance: one transposition counts as one edit, so "blod" is one step from "bold". */
function distance(a: string, b: string): number {
  const rows = a.length + 1
  const cols = b.length + 1
  const d: number[] = new Array(rows * cols)
  for (let i = 0; i < rows; i++) d[i * cols] = i
  for (let j = 0; j < cols; j++) d[j] = j
  for (let i = 1; i < rows; i++) {
    for (let j = 1; j < cols; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1
      let v = Math.min(
        d[(i - 1) * cols + j]! + 1,
        d[i * cols + j - 1]! + 1,
        d[(i - 1) * cols + j - 1]! + cost,
      )
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        v = Math.min(v, d[(i - 2) * cols + j - 2]! + 1)
      }
      d[i * cols + j] = v
    }
  }
  return d[rows * cols - 1]!
}

/** The closest candidate when the typo is small enough to be a typo; undefined when nothing is near. */
export function didYouMean(input: string, candidates: Iterable<string>): string | undefined {
  const needle = input.toLowerCase()
  if (!needle) return undefined
  const budget = Math.max(2, Math.floor(needle.length / 3))
  let best: { name: string; d: number } | undefined
  for (const name of candidates) {
    const d = distance(needle, name.toLowerCase())
    if (d > budget || d >= needle.length || (best && d >= best.d)) continue
    best = { name, d }
  }
  return best?.name
}

export function sheetNotFoundHints(name: string, sheets: readonly string[]): ErrorHints {
  const guess = didYouMean(name, sheets)
  return {
    reason: 'sheet_not_found',
    suggestion: guess
      ? `did you mean "${guess}"? (detail.sheets lists them all)`
      : 'use one of the names in detail.sheets',
  }
}
