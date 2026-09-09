/**
 * Borderless bank-statement layout (alpha feedback): every transaction
 * prints one value row (dates, description, amounts) followed by several
 * 1-unit description continuation lines (reference-number stacks). The
 * run-strength gate used to take its median over ALL run rows, so two
 * absorbed continuations collapsed it, absorption stopped, and the whole
 * region fell back to scrambled flow text. Strength is judged over the
 * anchor rows only.
 */
import { describe, expect, it } from 'vitest'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze/lines'
import { detectStreamTables } from '../src/analyze/stream'
import { splitIntoUnits } from '../src/analyze/units'
import type { PageShapes, PdfChar, TableBlock } from '../src/ir'
import { mkText } from './helpers/chars'

const shapesOf = (): PageShapes => ({ strokes: [], fills: [], ignoredPaths: 0 })
const unitsOf = (chars: PdfChar[]) => splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))

const cellText = (table: TableBlock, r: number, c: number): string =>
  table.rows[r]![c]!.blocks.map((b) =>
    b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(' '),
  ).join(' ')

/** right-aligned text: place so it ends at x1 (0.5 em per Latin glyph) */
function rightText(text: string, x1: number, y: number): PdfChar[] {
  return mkText(text, x1 - text.length * 5, { y }).chars
}

/** two transactions, each a 5-unit value row plus three continuation lines */
function statement(): PdfChar[] {
  const chars: PdfChar[] = []
  chars.push(...mkText('03 AUG', 40, { y: 730 }).chars)
  chars.push(...mkText('03 AUG', 95, { y: 730 }).chars)
  chars.push(...mkText('COMM/COMM IN LIEU', 150, { y: 730 }).chars)
  chars.push(...rightText('23.63', 430, 730))
  chars.push(...rightText('6,717,284.65', 610, 730))
  chars.push(...mkText('CT0039810817M001', 150, { y: 715 }).chars)
  chars.push(...mkText('SGTT260731808027', 150, { y: 700 }).chars)
  chars.push(...mkText('*MDSUSER 000003045', 150, { y: 685 }).chars)
  chars.push(...mkText('03 AUG', 40, { y: 670 }).chars)
  chars.push(...mkText('03 AUG', 95, { y: 670 }).chars)
  chars.push(...mkText('CALL A/C WDL', 150, { y: 670 }).chars)
  chars.push(...rightText('2,251,200.18', 430, 670))
  chars.push(...rightText('4,466,084.47', 610, 670))
  return chars
}

describe('detectStreamTables: bank statement with wrapped description stacks', () => {
  it('keeps the run alive across several absorbed continuation rows', () => {
    const { tables } = detectStreamTables(unitsOf(statement()), shapesOf())
    expect(tables).toHaveLength(1)
    const table = tables[0]!
    expect(table.confidence).toBeGreaterThanOrEqual(0.5)
    expect(table.rows).toHaveLength(5)
    expect(table.rows[0]).toHaveLength(5)
    // value rows land in their columns…
    expect(cellText(table, 0, 0)).toBe('03 AUG')
    expect(cellText(table, 0, 2)).toBe('COMM/COMM IN LIEU')
    expect(cellText(table, 0, 3)).toBe('23.63')
    expect(cellText(table, 0, 4)).toBe('6,717,284.65')
    expect(cellText(table, 4, 3)).toBe('2,251,200.18')
    // …and the reference stack stays inside the description column
    expect(cellText(table, 1, 2)).toBe('CT0039810817M001')
    expect(cellText(table, 2, 2)).toBe('SGTT260731808027')
    expect(cellText(table, 3, 2)).toBe('*MDSUSER 000003045')
    expect(cellText(table, 1, 0)).toBe('')
    expect(cellText(table, 1, 3)).toBe('')
  })
})
