export interface Span {
  start: number
  end: number
}

const WORD_SEGMENTER: Intl.Segmenter | null =
  typeof Intl !== 'undefined' && 'Segmenter' in Intl
    ? new Intl.Segmenter(undefined, { granularity: 'word' })
    : null

/** the word-like segment containing `off`, or the one ending exactly at it (caret right after a word) */
export function wordSegmentAt(text: string, off: number): Span | null {
  if (WORD_SEGMENTER) {
    let before: Span | null = null
    for (const seg of WORD_SEGMENTER.segment(text)) {
      const start = seg.index
      const end = seg.index + seg.segment.length
      if (start > off) break
      if (!seg.isWordLike) continue
      if (off < end) return { start, end }
      if (end === off) before = { start, end }
    }
    return before
  }
  const isWordChar = (ch: string) => /[\p{L}\p{N}_]/u.test(ch)
  let at = off
  if ((at >= text.length || !isWordChar(text[at]!)) && at > 0 && isWordChar(text[at - 1]!)) at--
  if (at >= text.length || !isWordChar(text[at]!)) return null
  let start = at
  while (start > 0 && isWordChar(text[start - 1]!)) start--
  let end = at + 1
  while (end < text.length && isWordChar(text[end]!)) end++
  return { start, end }
}

/** Word's word unit as a mouse gesture sees it: the word plus the spaces that follow it */
export function wordUnitAt(text: string, off: number): Span | null {
  const seg = wordSegmentAt(text, off)
  if (!seg) return null
  let end = seg.end
  while (end < text.length && text[end] === ' ') end++
  return { start: seg.start, end }
}
