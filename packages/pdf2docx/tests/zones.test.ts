/** Rule-separated side-by-side zone detection (P22 A): hand-built units, no wasm. */
import { describe, expect, it } from 'vitest'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze/lines'
import { splitIntoUnits } from '../src/analyze/units'
import { detectRuleSeparatedZones } from '../src/analyze/zones'
import type { Fill, PageShapes, PdfChar, Stroke } from '../src/ir'
import { mkText } from './helpers/chars'

const shapesOf = (strokes: Stroke[] = [], fills: Fill[] = []): PageShapes => ({
  strokes,
  fills,
  ignoredPaths: 0,
})

const vStroke = (x: number, y0: number, y1: number): Stroke => ({
  box: { x0: x - 0.45, x1: x + 0.45, y0, y1 },
  orientation: 'v',
  widthPt: 0.9,
  color: '000000',
})

const unitsOf = (chars: PdfChar[]) => splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))

const PAGE = { pageWidthPt: 612, pageHeightPt: 792 }

/** court-caption shape: left party stack, right title stack, rule at x=300 */
function captionChars(): PdfChar[] {
  return [
    ...mkText('DOCTOR JOHNS INC an Iowa', 80, { y: 560 }).chars,
    ...mkText('Corporation,', 80, { y: 540 }).chars,
    ...mkText('Plaintiff,', 150, { y: 515 }).chars,
    ...mkText('No. C 03-4121-MWB', 360, { y: 515 }).chars,
    ...mkText('vs.', 80, { y: 490 }).chars,
    ...mkText('ORDER REGARDING SANCTIONS', 320, { y: 490 }).chars,
    ...mkText('CITY OF SIOUX CITY IOWA and', 80, { y: 465 }).chars,
    ...mkText('OF RELEVANT RECORDS', 340, { y: 460 }).chars,
    ...mkText('Defendants.', 150, { y: 430 }).chars,
  ]
}

/** rule broken into segments, like PDF double rules drawn piecewise */
const captionRule = (): Stroke[] => [
  vStroke(300, 420, 480),
  vStroke(300, 480, 570),
  vStroke(302, 420, 500),
  vStroke(302, 500, 570),
]

describe('detectRuleSeparatedZones', () => {
  it('claims a rule-separated caption as one borderless 1×2 table', () => {
    const units = unitsOf(captionChars())
    const res = detectRuleSeparatedZones(units, shapesOf(captionRule()), [], PAGE)
    expect(res.tables).toHaveLength(1)
    const t = res.tables[0]!
    expect(t.rows).toHaveLength(1)
    expect(t.rows[0]).toHaveLength(2)
    expect(t.confidence).toBeGreaterThan(0.5)
    expect(t.sepRule).toBe('double')
    expect(res.consumedStrokes.size).toBe(4)
    // all caption units consumed, none left in the flow
    expect(res.remainingUnits).toHaveLength(0)
    const cellText = (c: number): string =>
      t.rows[0]![c]!.blocks.map((b) =>
        b.lines.map((l) => l.spans.map((s) => s.text).join('')).join('|'),
      ).join('|')
    expect(cellText(0)).toContain('Plaintiff')
    expect(cellText(0)).not.toContain('ORDER')
    expect(cellText(1)).toContain('ORDER REGARDING SANCTIONS')
    expect(cellText(1)).not.toContain('Defendants')
    // display lines are re-sorted top→down inside each cell
    expect(cellText(0).indexOf('Corporation')).toBeLessThan(cellText(0).indexOf('Defendants'))
  })

  it('reads a single drawn line as a single rule', () => {
    const units = unitsOf(captionChars())
    const res = detectRuleSeparatedZones(units, shapesOf([vStroke(300, 420, 570)]), [], PAGE)
    expect(res.tables).toHaveLength(1)
    expect(res.tables[0]!.sepRule).toBe('single')
  })

  it('does not fire without a drawn rule (gutter alone is not evidence)', () => {
    const units = unitsOf(captionChars())
    const res = detectRuleSeparatedZones(units, shapesOf([]), [], PAGE)
    expect(res.tables).toHaveLength(0)
    expect(res.remainingUnits).toHaveLength(units.length)
  })

  it('aborts when a unit crosses the rule', () => {
    const chars = [
      ...captionChars(),
      ...mkText('A crossing line of text here', 200, { y: 505 }).chars,
    ]
    const units = unitsOf(chars)
    const res = detectRuleSeparatedZones(units, shapesOf(captionRule()), [], PAGE)
    expect(res.tables).toHaveLength(0)
  })

  it('ignores short rules (a few lines tall is not a zone separator)', () => {
    const units = unitsOf(captionChars())
    const res = detectRuleSeparatedZones(units, shapesOf([vStroke(300, 500, 530)]), [], PAGE)
    expect(res.tables).toHaveLength(0)
  })

  it('ignores rules inside claimed (lattice) regions', () => {
    const units = unitsOf(captionChars())
    const res = detectRuleSeparatedZones(
      units,
      shapesOf(captionRule()),
      [{ x0: 60, x1: 560, y0: 400, y1: 600 }],
      PAGE,
    )
    expect(res.tables).toHaveLength(0)
  })

  it('does not fire on landscape pages', () => {
    const units = unitsOf(captionChars())
    const res = detectRuleSeparatedZones(units, shapesOf(captionRule()), [], {
      pageWidthPt: 792,
      pageHeightPt: 612,
    })
    expect(res.tables).toHaveLength(0)
  })

  it('requires stacks on BOTH sides of the rule', () => {
    const chars = [
      ...mkText('Left only line one', 80, { y: 560 }).chars,
      ...mkText('Left only line two', 80, { y: 530 }).chars,
      ...mkText('Left only line three', 80, { y: 500 }).chars,
    ]
    const units = unitsOf(chars)
    const res = detectRuleSeparatedZones(units, shapesOf(captionRule()), [], PAGE)
    expect(res.tables).toHaveLength(0)
  })
})
