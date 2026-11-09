import type { RecentEntry } from '../../shared/home-api'

/** every path between `anchor` and `target` (inclusive) in list order; just the target without an anchor */
export function rangeSelect(
  paths: readonly string[],
  anchor: string | null,
  target: string,
): string[] {
  const to = paths.indexOf(target)
  if (to < 0) return []
  const from = anchor === null ? -1 : paths.indexOf(anchor)
  if (from < 0) return [target]
  const [lo, hi] = from <= to ? [from, to] : [to, from]
  return paths.slice(lo, hi + 1)
}

/** keyboard row navigation; -1 current means "nothing focused yet" */
export function nextRowIndex(
  key: 'ArrowDown' | 'ArrowUp' | 'Home' | 'End',
  current: number,
  count: number,
): number {
  if (count <= 0) return -1
  switch (key) {
    case 'Home':
      return 0
    case 'End':
      return count - 1
    case 'ArrowDown':
      return Math.min(count - 1, current + 1)
    case 'ArrowUp':
      return current <= 0 ? 0 : current - 1
  }
}

/** true when the key press belongs to a text field and list shortcuts must stay out of the way */
export function isEditableTarget(target: EventTarget | null): boolean {
  const el = target as { tagName?: string; isContentEditable?: boolean; type?: string } | null
  if (!el || typeof el.tagName !== 'string') return false
  const tag = el.tagName.toLowerCase()
  if (tag === 'textarea' || tag === 'select') return true
  if (tag === 'input') return el.type !== 'checkbox' && el.type !== 'button'
  return el.isContentEditable === true
}

export type FileSort = 'recent' | 'oldest' | 'nameAsc' | 'nameDesc' | 'sizeDesc' | 'sizeAsc'

/** list order for the file table; `recent` is the source order (newest first) */
export function sortEntries(entries: readonly RecentEntry[], sort: FileSort): RecentEntry[] {
  const list = [...entries]
  const byName = (a: RecentEntry, b: RecentEntry): number =>
    a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' })
  switch (sort) {
    case 'recent':
      return list
    case 'oldest':
      return list.reverse()
    case 'nameAsc':
      return list.sort(byName)
    case 'nameDesc':
      return list.sort((a, b) => byName(b, a))
    case 'sizeDesc':
      return list.sort((a, b) => b.sizeBytes - a.sizeBytes)
    case 'sizeAsc':
      return list.sort((a, b) => a.sizeBytes - b.sizeBytes)
  }
}

/** only these open in an editor tab; anything else dropped on the window is skipped with a hint */
const OPENABLE = new Set([
  'docx',
  'doc',
  'xlsx',
  'xls',
  'csv',
  'tsv',
  'pptx',
  'ppt',
  'pdf',
  'md',
  'markdown',
  'txt',
  'json',
  'xml',
  'yaml',
  'yml',
  'sql',
])

export function isOpenablePath(path: string): boolean {
  const dot = path.lastIndexOf('.')
  return dot >= 0 && OPENABLE.has(path.slice(dot + 1).toLowerCase())
}
