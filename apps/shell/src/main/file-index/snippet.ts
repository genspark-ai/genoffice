import { findRanges, findTextRanges, foldText, type TextMark } from '../../shared/text-marks'

export type SnippetPart = TextMark

const WINDOW = 150
const LEAD = 40

/**
 * A ~150-char excerpt around the first matched needle with every needle
 * occurrence inside it flagged. Returns null when no needle occurs.
 */
export function buildSnippet(text: string, needles: readonly string[]): SnippetPart[] | null {
  if (needles.every((n) => !n) || !text) return null
  const compact = text.replace(/\s+/g, ' ')
  const ranges = findTextRanges(compact, needles)
  if (ranges.length === 0) return null
  const start = Math.max(0, ranges[0]![0] - LEAD)
  const end = Math.min(compact.length, start + WINDOW)
  const parts: SnippetPart[] = []
  let cursor = start
  if (start > 0) parts.push({ text: '…', hit: false })
  for (const [s, e] of ranges) {
    if (s >= end) break
    if (s > cursor) parts.push({ text: compact.slice(cursor, s), hit: false })
    const stop = Math.min(e, end)
    parts.push({ text: compact.slice(Math.max(s, cursor), stop), hit: true })
    cursor = stop
  }
  if (cursor < end) parts.push({ text: compact.slice(cursor, end), hit: false })
  if (end < compact.length) parts.push({ text: '…', hit: false })
  return parts
}

/** up to `chars` of text starting a little before the first needle hit, or the head of the text */
export function excerpt(text: string, needles: readonly string[], chars: number): string {
  const compact = text.replace(/\s+/g, ' ').trim()
  const ranges = findTextRanges(compact, needles)
  const start = ranges.length ? Math.max(0, ranges[0]![0] - Math.floor(chars / 4)) : 0
  return compact.slice(start, start + chars)
}

/** true when any needle occurs in the (folded) text */
export function containsAny(text: string, needles: readonly string[]): boolean {
  return findRanges(foldText(text), needles.map(foldText).filter(Boolean)).length > 0
}
