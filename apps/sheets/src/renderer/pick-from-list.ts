/**
 * Excel's Pick From Drop-down List (Alt+Down): the unique entries of the
 * contiguous non-empty block the active cell sits in (or borders, when the
 * cell itself is empty), sorted the way Excel lists them.
 */
import type { ICellData } from '@univerjs/core'

export interface PickListEntry {
  readonly display: string
  readonly cell: ICellData | null
}

export const PICK_LIST_SCAN_LIMIT = 10_000

const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' })

function isBlank(entry: PickListEntry | null): boolean {
  return !entry || entry.display.trim() === ''
}

export function collectPickListValues(
  read: (row: number) => PickListEntry | null,
  row: number,
  rowCount: number,
): PickListEntry[] {
  const block: PickListEntry[] = []
  for (let r = row - 1, steps = 0; r >= 0 && steps < PICK_LIST_SCAN_LIMIT; r -= 1, steps += 1) {
    const entry = read(r)
    if (isBlank(entry)) break
    block.push(entry!)
  }
  const self = read(row)
  if (!isBlank(self)) block.push(self!)
  for (
    let r = row + 1, steps = 0;
    r < rowCount && steps < PICK_LIST_SCAN_LIMIT;
    r += 1, steps += 1
  ) {
    const entry = read(r)
    if (isBlank(entry)) break
    block.push(entry!)
  }
  const seen = new Set<string>()
  const unique: PickListEntry[] = []
  for (const entry of block) {
    const key = entry.display.trim().toLocaleLowerCase()
    if (seen.has(key)) continue
    seen.add(key)
    unique.push(entry)
  }
  return unique.sort((a, b) => collator.compare(a.display, b.display))
}

/** What Univer's value filter compares against (`extractPureTextFromCell`). */
export function filterCriteriaForText(text: string): { blank?: true; filters?: string[] } {
  return text === '' ? { blank: true } : { filters: [text] }
}
