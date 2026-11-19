/// Remove Duplicates over a value matrix: keeps the first
/// occurrence of each row, compares text case-insensitively, and leaves an
/// optional header row in place.

export type DedupeValue = string | number | boolean | null

export function dedupeRows(
  rows: ReadonlyArray<ReadonlyArray<DedupeValue>>,
  hasHeader: boolean,
): { rows: DedupeValue[][]; removed: number } {
  const kept: DedupeValue[][] = []
  const seen = new Set<string>()
  let removed = 0
  for (const [index, row] of rows.entries()) {
    if (hasHeader && index === 0) {
      kept.push([...row])
      continue
    }
    const key = JSON.stringify(
      row.map((value) => (typeof value === 'string' ? value.toLowerCase() : value)),
    )
    if (seen.has(key)) {
      removed += 1
      continue
    }
    seen.add(key)
    kept.push([...row])
  }
  return { rows: kept, removed }
}
