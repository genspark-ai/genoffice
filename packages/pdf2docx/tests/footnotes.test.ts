/** Footnote detection unit tests: hand-built chars + separator strokes (no wasm). */
import { describe, expect, it } from 'vitest'
import { detectFootnotes } from '../src/analyze/footnotes'
import type { PageShapes, PdfChar, Stroke } from '../src/ir'
import { mkChar, mkText } from './helpers/chars'

const W = 612
const H = 792

const sepStroke = (y: number, x0 = 64, len = 144): Stroke => ({
  box: { x0, x1: x0 + len, y0: y - 0.1, y1: y + 0.1 },
  orientation: 'h',
  widthPt: 0.5,
  color: '000000',
})

const shapesOf = (strokes: Stroke[]): PageShapes => ({ strokes, fills: [], ignoredPaths: 0 })

/** body paragraph + anchored superscript + separator + marker line + note text */
function notePage(): { chars: PdfChar[]; shapes: PageShapes } {
  const bodyLine = mkText('The purpose of this system', 72, { y: 700, fontSize: 12 })
  const sup = mkChar('1', bodyLine.endX + 1, { y: 705, fontSize: 7 })
  const marker = mkChar('1', 64, { y: 390, fontSize: 7 })
  const note = mkText('State the final goal here.', 72, { y: 382, fontSize: 9 }).chars
  return {
    chars: [...bodyLine.chars, sup, marker, ...note],
    shapes: shapesOf([sepStroke(400)]),
  }
}

const noteText = (f: { blocks: Array<{ lines: Array<{ spans: Array<{ text: string }> }> }> }) =>
  f.blocks.map((b) => b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(' ')).join(' ')

describe('detectFootnotes', () => {
  it('lifts a marker-line note off the page and collapses the body superscript', () => {
    const { chars, shapes } = notePage()
    const { bodyChars, footnotes } = detectFootnotes(chars, shapes, 0, W, H)
    expect(footnotes).toHaveLength(1)
    expect(noteText(footnotes[0]!)).toBe('State the final goal here.')
    // note lines gone from the body
    expect(bodyChars.some((c) => c.text === 'g')).toBe(false)
    // the superscript digit became a zero-text anchor
    const anchors = bodyChars.filter((c) => c.noteRef !== undefined)
    expect(anchors).toHaveLength(1)
    expect(anchors[0]!.text).toBe('')
    expect(anchors[0]!.noteRef).toBe(footnotes[0]!.id)
  })

  it('parses inline markers ("1 Note text…") and roman marker lines', () => {
    const bodyLine = mkText('Body content up here', 72, { y: 700, fontSize: 12 })
    const inline = [
      mkChar('1', 64, { y: 388, fontSize: 7 }),
      ...mkText(' First note.', 70, { y: 385, fontSize: 10 }).chars,
    ]
    const { footnotes } = detectFootnotes(
      [...bodyLine.chars, ...inline],
      shapesOf([sepStroke(400)]),
      0,
      W,
      H,
    )
    expect(footnotes).toHaveLength(1)
    expect(noteText(footnotes[0]!)).toContain('First note.')

    const roman = [
      mkChar('i', 64, { y: 390, fontSize: 6 }),
      ...mkText('Endnote body.', 72, { y: 382, fontSize: 10 }).chars,
    ]
    const res = detectFootnotes([...bodyLine.chars, ...roman], shapesOf([sepStroke(400)]), 0, W, H)
    expect(res.footnotes).toHaveLength(1)
  })

  it('synthesizes an end-of-body anchor when the superscript is elsewhere', () => {
    const bodyLine = mkText('Section content without a marker', 72, { y: 700, fontSize: 12 })
    const marker = mkChar('7', 64, { y: 390, fontSize: 7 })
    const note = mkText('Cross-page note.', 72, { y: 382, fontSize: 9 }).chars
    const { bodyChars, footnotes } = detectFootnotes(
      [...bodyLine.chars, marker, ...note],
      shapesOf([sepStroke(400)]),
      2,
      W,
      H,
    )
    expect(footnotes).toHaveLength(1)
    const anchors = bodyChars.filter((c) => c.noteRef !== undefined)
    expect(anchors).toHaveLength(1)
    expect(anchors[0]!.isGenerated).toBe(true)
  })

  it('rejects pages without a separator or with body text below the rule', () => {
    const { chars } = notePage()
    expect(detectFootnotes(chars, shapesOf([]), 0, W, H).footnotes).toHaveLength(0)

    const bodyLine = mkText('Heading over a rule', 72, { y: 700, fontSize: 12 })
    const below = mkText('Ordinary body text below the rule.', 72, { y: 380, fontSize: 12 })
    const res = detectFootnotes(
      [...bodyLine.chars, ...below.chars],
      shapesOf([sepStroke(400)]),
      0,
      W,
      H,
    )
    expect(res.footnotes).toHaveLength(0)
    expect(res.bodyChars).toHaveLength(bodyLine.chars.length + below.chars.length)
  })
})
