import type { SectionSettings } from '@genoffice/docx-engine'

/** how a page's footnote area is laid out: its columns and the gap between them (px) */
export interface NoteColumns {
  count: number
  gapPx: number
  /** a line between the columns: the section's own, where the area matches it */
  rule?: true
}

/** a note as the heights of its line boxes (px); one entry = a note that cannot be divided */
export type NoteLines = readonly number[]

const twipsToPx = (twips: number) => (twips / 1440) * 96
/** Word's gap between footnote columns a section's own columns do not give: half an inch */
const DEFAULT_GAP_TWIPS = 720
export const MAX_NOTE_COLUMNS = 4

type ColumnsOf = Pick<SectionSettings, 'columns' | 'colSpace' | 'colSep' | 'footnoteColumns'>

/**
 * Columns of the footnote area of a page whose first note is in the section: its
 * w15:footnoteColumns, else as many as it has text columns. Documents laid out as
 * Word 2010 or older (compatibilityMode below 15) keep one column.
 */
export function noteColumnsOf(section: ColumnsOf, compatibilityMode: number | undefined): number {
  if ((compatibilityMode ?? 0) < 15) return 1
  const own = section.footnoteColumns
  if (own !== undefined) return Math.max(1, Math.min(own, MAX_NOTE_COLUMNS))
  return Math.max(1, section.columns)
}

/** the footnote columns a section gives; null where the area stays one column */
export function noteColumnLayout(
  section: ColumnsOf,
  compatibilityMode: number | undefined,
): NoteColumns | null {
  const count = noteColumnsOf(section, compatibilityMode)
  if (count <= 1) return null
  const matched = section.footnoteColumns === undefined
  return {
    count,
    gapPx: twipsToPx((matched && section.colSpace) || DEFAULT_GAP_TWIPS),
    ...(matched && section.colSep ? { rule: true as const } : {}),
  }
}

/** width of one footnote column in a content area `contentPx` wide */
export function noteColumnWidth(contentPx: number, cols: NoteColumns | null): number {
  if (!cols) return contentPx
  return Math.max((contentPx - cols.gapPx * (cols.count - 1)) / cols.count, 1)
}

/**
 * The line boxes of a note: the first carries what is left over (its raised number
 * makes that line taller). A note far from whole lines is one piece.
 */
export function noteLines(height: number, lineHeight: number | undefined): number[] {
  if (!lineHeight || lineHeight <= 0 || height <= 0) return [Math.max(height, 0)]
  const n = Math.round(height / lineHeight)
  const over = height - n * lineHeight
  if (n < 1 || over < -1 || over > lineHeight * 0.4) return [height]
  return Array.from({ length: n }, (_, i) => (i === 0 ? lineHeight + over : lineHeight))
}

const EPS = 0.01

/**
 * Fill `count` columns `height` tall, one after the other. A note continues in the
 * next column with at least two lines on either side (widow / orphan control).
 * Returns the tallest column, or the least height at which more would fit.
 */
function fill(
  notes: readonly NoteLines[],
  count: number,
  height: number,
): { fits: true; used: number } | { fits: false; need: number } {
  let col = 0
  let used = 0
  let tallest = 0
  let need = Infinity
  const turn = (blocked: number): boolean => {
    need = Math.min(need, used + blocked)
    tallest = Math.max(tallest, used)
    used = 0
    return ++col < count
  }
  for (const lines of notes) {
    let at = 0
    while (at < lines.length) {
      const rest = lines.length - at
      let take = 0
      let h = 0
      while (take < rest && used + h + lines[at + take] <= height + EPS) h += lines[at + take++]
      if (take < rest) {
        if (rest - take === 1) take -= 1
        if (take < 2) take = 0
      }
      if (take === 0) {
        // the least this note could put here: all of it when it is too short to divide
        const unit = rest <= 3 ? rest : 2
        const blocked = lines.slice(at, at + unit).reduce((a, b) => a + b, 0)
        if (used <= EPS) return { fits: false, need: blocked }
        if (!turn(blocked)) return { fits: false, need }
        continue
      }
      for (let k = 0; k < take; k++) used += lines[at + k]
      at += take
    }
  }
  return { fits: true, used: Math.max(tallest, used) }
}

/** Height of a footnote area: the least at which the notes, filled column after column, fit */
export function noteAreaHeight(notes: readonly NoteLines[], count: number): number {
  const total = notes.reduce((sum, lines) => sum + lines.reduce((a, b) => a + b, 0), 0)
  if (count <= 1 || total <= 0) return total
  let height = total / count
  for (let guard = 0; guard < 10000; guard++) {
    const r = fill(notes, count, height)
    if (r.fits) return r.used
    if (!(r.need > height + EPS)) break
    height = r.need
  }
  return total
}

/**
 * What each note adds to an area that opens with the first of them. Past `limit`
 * (no page holds a taller area) the rest count one under the other.
 */
export function noteAreaIncrements(
  notes: readonly NoteLines[],
  count: number,
  limit = Infinity,
): number[] {
  const out: number[] = []
  let prev = 0
  let k = 0
  for (; k < notes.length && prev <= limit; k++) {
    // widow / orphan control can make a longer list a shade shorter: an area never shrinks
    const h = Math.max(prev, noteAreaHeight(notes.slice(0, k + 1), count))
    out.push(h - prev)
    prev = h
  }
  for (; k < notes.length; k++) out.push(notes[k].reduce((a, b) => a + b, 0))
  return out
}

/** height of a page's footnote area without its separator */
export function noteAreaPx(
  items: ReadonlyArray<{ height: number; lineHeightPx?: number }>,
  cols: NoteColumns | null,
): number {
  if (!cols) return items.reduce((s, n) => s + n.height, 0)
  return noteAreaHeight(
    items.map((n) => noteLines(n.height, n.lineHeightPx)),
    cols.count,
  )
}

/** class of the box that lays a footnote area's rows out in columns */
export const noteColumnsClass = (cols: NoteColumns): string =>
  cols.rule ? 'note-columns note-columns-rule' : 'note-columns'

/** Inline style of that box: filled to the height the pagination reserved, plus a pixel for rounding */
export function noteColumnsStyle(
  items: ReadonlyArray<{ height: number; lineHeightPx?: number }>,
  cols: NoteColumns,
): { columnCount: string; columnGap: string; height: string } {
  return {
    columnCount: String(cols.count),
    columnGap: `${cols.gapPx}px`,
    height: `${Math.ceil(noteAreaPx(items, cols)) + 1}px`,
  }
}
