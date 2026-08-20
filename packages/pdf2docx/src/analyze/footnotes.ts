/**
 * Footnote detection (P6). A rendered footnote block is body text to the
 * naive pipeline, so every page's notes leaked into the rebuilt document
 * flow. The layout is distinctive: a short thin separator rule at the
 * content's left edge in the lower page half, below it note entries each
 * opened by a small (superscript-sized) bare number, and matching raised
 * superscript digits in the body text above.
 *
 * Detected notes are lifted off the page into IrPage.footnotes; the body
 * superscript marker collapses into a zero-text anchor char whose
 * `noteRef` the rebuild layer turns into a real w:footnoteReference.
 * Anything that does not parse cleanly (a line below the rule before any
 * marker, non-consecutive numbers, an unmatched anchor) stays in the body —
 * miss rather than misfire.
 */
import { median } from '../geometry'
import type { FootnoteIR, PageShapes, PdfChar } from '../ir'
import { groupIntoBlocks } from './blocks'
import { analyzeChars } from './chars'
import type { RawLine } from './lines'
import { groupIntoLines, isSpaceCode } from './lines'

/**
 * separator rule geometry: thin, short, at the content left edge. Endnotes
 * rendered at a section's end put the rule anywhere on the page (ja-form
 * documents place one per yoshiki form page), so only the extreme page top is out.
 */
const SEP_MAX_WIDTH_PT = 2
const SEP_MIN_LEN_PT = 36
const SEP_MAX_LEN_RATIO = 0.5
const SEP_LEFT_TOL_RATIO = 0.15
const SEP_MAX_Y_RATIO = 0.9
/** marker digits are superscript-sized against their note text / neighbours */
const MARKER_FONT_RATIO = 0.85
const ANCHOR_FONT_RATIO = 0.85
/** body anchors sit visibly above their neighbour's baseline (share of font size) */
const ANCHOR_RAISE_RATIO = 0.12
/** sanity bounds: markers count up one by one from a plausible start */
const MARKER_MAX_START = 150
const MARKER_MAX_NOTES = 30

const isVisible = (c: PdfChar): boolean => !isSpaceCode(c.code) && c.code > 0x1f
const isDigit = (c: PdfChar): boolean => /^[0-9０-９]$/.test(c.text)
/** fullwidth digits read as their ASCII value (ja documents number notes ７, ８…) */
const digitValue = (chars: readonly PdfChar[]): number =>
  parseInt(
    chars
      .map((c) => String.fromCharCode(c.text.charCodeAt(0) - (c.text >= '０' ? 0xfee0 : 0)))
      .join(''),
    10,
  )

export interface DetectedFootnotes {
  /** page chars with note lines removed and anchors collapsed */
  bodyChars: PdfChar[]
  footnotes: FootnoteIR[]
}

const NONE = (chars: readonly PdfChar[]): DetectedFootnotes => ({
  bodyChars: [...chars],
  footnotes: [],
})

interface ParsedNote {
  num: number
  chars: PdfChar[]
}

/** lowercase roman numeral → value (endnote markers: i, ii, … xxx); null otherwise */
function romanValue(text: string): number | null {
  if (!/^[ivx]{1,6}$/.test(text)) return null
  const one: Record<string, number> = { i: 1, v: 5, x: 10 }
  let total = 0
  for (let i = 0; i < text.length; i++) {
    const cur = one[text[i]!]!
    const next = one[text[i + 1] ?? ''] ?? 0
    total += cur < next ? -cur : cur
  }
  return total >= 1 && total <= 40 ? total : null
}

/** a line that is ONLY a small bare number (the separated marker of a note) */
function markerLineValue(line: RawLine, bodyFont: number): number | null {
  const visible = line.chars.filter(isVisible)
  if (visible.length === 0 || visible.length > 3) return null
  const font = Math.max(...visible.map((c) => c.fontSize))
  if (font > MARKER_FONT_RATIO * bodyFont) return null
  if (visible.every(isDigit)) return digitValue(visible)
  return romanValue(visible.map((c) => c.text).join(''))
}

/** "1 Note text…" on one line: leading small digits, then normal-size text */
function inlineMarker(line: RawLine): { num: number; rest: PdfChar[] } | null {
  const chars = line.chars
  let i = 0
  while (i < chars.length && !isVisible(chars[i]!)) i++
  const digits: PdfChar[] = []
  while (i < chars.length && isDigit(chars[i]!) && digits.length <= 3) {
    digits.push(chars[i]!)
    i++
  }
  if (digits.length === 0 || digits.length > 3) return null
  const rest = chars.slice(i)
  const restVisible = rest.filter(isVisible)
  if (restVisible.length === 0) return null
  const restFont = median(restVisible.map((c) => c.fontSize))
  if (Math.max(...digits.map((c) => c.fontSize)) > MARKER_FONT_RATIO * restFont) return null
  return { num: digitValue(digits), rest }
}

/** note numbering steps by one, repeats (repeated template) or restarts at 1 */
const nextNumOk = (prev: number, next: number): boolean =>
  next === prev + 1 || next === prev || next === 1

/** parse every line under the separator into consecutive numbered notes */
function parseNotes(below: RawLine[], bodyFont: number): ParsedNote[] | null {
  const notes: ParsedNote[] = []
  let current: ParsedNote | null = null
  for (const line of below) {
    const markerValue = markerLineValue(line, bodyFont)
    if (markerValue !== null) {
      if (current) notes.push(current)
      current = { num: markerValue, chars: [] }
      continue
    }
    const inline = inlineMarker(line)
    if (inline && (current === null || nextNumOk(current.num, inline.num))) {
      if (current) notes.push(current)
      current = { num: inline.num, chars: [...inline.rest] }
      continue
    }
    if (current === null) return null // content below the rule before any marker
    current.chars.push(...line.chars)
  }
  if (current) notes.push(current)

  if (notes.length === 0 || notes.length > MARKER_MAX_NOTES) return null
  if (notes[0]!.num < 1 || notes[0]!.num > MARKER_MAX_START) return null
  for (const [i, note] of notes.entries()) {
    if (i > 0 && !nextNumOk(notes[i - 1]!.num, note.num)) return null
    if (!note.chars.some(isVisible)) return null
  }
  return notes
}

/**
 * Find the superscript digit run reading `num` in the body chars: digits far
 * smaller than their neighbour and raised off its baseline. Returns char
 * indexes into `body`, or null.
 */
function findAnchor(body: readonly PdfChar[], num: number): number[] | null {
  const wanted = String(num)
  for (let i = 0; i < body.length; i++) {
    if (body[i]!.noteRef !== undefined) continue
    const idx: number[] = []
    for (let j = i; j < body.length && idx.length < wanted.length; j++) {
      const c = body[j]!
      if (!isVisible(c)) break
      if (!isDigit(c) || c.noteRef !== undefined) break
      idx.push(j)
    }
    if (idx.length === 0 || String(digitValue(idx.map((at) => body[at]!))) !== wanted) continue
    // neighbour = nearest visible non-digit char before (after as fallback)
    let nb: PdfChar | undefined
    for (let j = i - 1; j >= 0; j--) {
      const c = body[j]!
      if (isVisible(c) && !idx.includes(j)) {
        nb = c
        break
      }
    }
    if (!nb) {
      const after = body[idx[idx.length - 1]! + 1]
      if (after && isVisible(after)) nb = after
    }
    if (!nb) continue
    const c0 = body[i]!
    const raised = c0.originY >= nb.originY + ANCHOR_RAISE_RATIO * nb.fontSize
    const small = c0.fontSize <= ANCHOR_FONT_RATIO * nb.fontSize
    if (raised && small) return idx
  }
  return null
}

/** Full footnote pass over one page. `pageIndex` keys the document-unique ids. */
export function detectFootnotes(
  chars: readonly PdfChar[],
  shapes: PageShapes,
  pageIndex: number,
  widthPt: number,
  heightPt: number,
): DetectedFootnotes {
  const visible = chars.filter(isVisible)
  if (visible.length === 0 || shapes.strokes.length === 0) return NONE(chars)
  const contentLeft = Math.min(...visible.map((c) => c.box.x0))
  const bodyFont = median(visible.map((c) => c.fontSize)) || 12

  // candidate separator rules, topmost first — body text below a decorative
  // underline never parses as notes, so the scan falls through to the real one
  const separators = shapes.strokes
    .filter((s) => {
      if (s.orientation !== 'h' || s.widthPt > SEP_MAX_WIDTH_PT) return false
      const len = s.box.x1 - s.box.x0
      if (len < SEP_MIN_LEN_PT || len > SEP_MAX_LEN_RATIO * widthPt) return false
      if (s.box.x0 > contentLeft + SEP_LEFT_TOL_RATIO * widthPt) return false
      return (s.box.y0 + s.box.y1) / 2 < SEP_MAX_Y_RATIO * heightPt
    })
    .sort((a, b) => b.box.y1 - a.box.y1)

  const lines = groupIntoLines(chars)
  for (const sep of separators) {
    const sepY = (sep.box.y0 + sep.box.y1) / 2
    const below = lines.filter((l) => l.baseline < sepY).sort((a, b) => b.baseline - a.baseline)
    if (below.length === 0) continue
    const notes = parseNotes(below, bodyFont)
    if (!notes) continue

    // resolve anchors against the body (everything except the note lines)
    const noteChars = new Set<PdfChar>()
    for (const line of below) for (const c of line.chars) noteChars.add(c)
    const body = chars.filter((c) => !noteChars.has(c))
    const lastVisibleBody = [...body].reverse().find(isVisible)
    if (!lastVisibleBody) return NONE(chars) // notes-only page: leave as body text

    const footnotes: FootnoteIR[] = []
    const dropFromBody = new Set<PdfChar>()
    const syntheticAnchors: PdfChar[] = []
    for (const [seq, note] of notes.entries()) {
      // sequence-keyed (marker numbers repeat when a template repeats)
      const id = String((pageIndex + 1) * 100 + seq + 1)
      const anchor = findAnchor(body, note.num)
      if (anchor) {
        const first = body[anchor[0]]!
        first.noteRef = id
        first.text = ''
        for (const at of anchor.slice(1)) dropFromBody.add(body[at]!)
      } else {
        // the superscript sits on an earlier page (endnotes collect at their
        // section's end): reference the note after the page's last body char
        syntheticAnchors.push({
          ...lastVisibleBody,
          text: '',
          noteRef: id,
          box: { ...lastVisibleBody.box, x0: lastVisibleBody.box.x1 },
          looseBox: { ...lastVisibleBody.looseBox, x0: lastVisibleBody.looseBox.x1 },
          originX: lastVisibleBody.looseBox.x1,
          isGenerated: true,
        })
      }
      footnotes.push({ id, blocks: groupIntoBlocks(analyzeChars(note.chars)) })
    }

    const bodyChars: PdfChar[] = []
    for (const c of chars) {
      if (noteChars.has(c) || dropFromBody.has(c)) continue
      bodyChars.push(c)
      if (c === lastVisibleBody) bodyChars.push(...syntheticAnchors)
    }
    // the matched separator is consumed with its notes (P12 D): Word draws
    // its own rule above the footnote area, and the leftover stroke rebuilt
    // as a decor rule claims spacing down to the page bottom that the
    // footnote area now occupies — on flush pages that hairline spills an
    // entirely blank page
    const sepIdx = shapes.strokes.indexOf(sep)
    if (sepIdx >= 0) shapes.strokes.splice(sepIdx, 1)
    return { bodyChars, footnotes }
  }
  return NONE(chars)
}
