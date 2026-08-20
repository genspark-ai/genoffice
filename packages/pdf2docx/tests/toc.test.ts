/** TOC-entry detection unit tests: dot leaders + leaderless rows (no wasm). */
import { describe, expect, it } from 'vitest'
import { groupIntoBlocks } from '../src/analyze/blocks'
import { analyzeChars } from '../src/analyze/chars'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze/lines'
import { detectTocBlocks, detectTocRows } from '../src/analyze/toc'
import { splitIntoUnits } from '../src/analyze/units'
import type { PdfChar } from '../src/ir'
import { mkChar, mkText } from './helpers/chars'

const unitsOf = (chars: PdfChar[]) => splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))

/** "Title ........ 42" laid out with real dot glyphs */
function dotLeaderChars(title: string, page: string, y: number): PdfChar[] {
  const t = mkText(title, 72, { y })
  const chars = [...t.chars]
  for (let x = Math.ceil(t.endX) + 6; x <= 480; x += 6) chars.push(mkChar('.', x, { y, width: 2 }))
  chars.push(...mkText(page, 500, { y }).chars)
  return chars
}

describe('detectTocBlocks (dot leaders)', () => {
  it('splits leader lines into TOC entries, title only, dots dropped', () => {
    const chars = [
      ...dotLeaderChars('Introduction', '3', 700),
      ...dotLeaderChars('Background', '7', 686),
    ]
    const blocks = detectTocBlocks(groupIntoBlocks(analyzeChars(chars)))
    expect(blocks).toHaveLength(2)
    expect(blocks[0]!.tocEntry).toEqual({ level: 1, pageNumber: '3' })
    expect(blocks[0]!.lines[0]!.spans.map((s) => s.text).join('')).toBe('Introduction')
    expect(blocks[1]!.tocEntry?.pageNumber).toBe('7')
  })

  it('leaves ordinary paragraphs untouched', () => {
    const blocks = groupIntoBlocks(
      analyzeChars(mkText('Just a sentence with no leader.', 72).chars),
    )
    expect(detectTocBlocks(blocks)).toBe(blocks)
  })
})

describe('detectTocRows (leaderless entries)', () => {
  const entryRow = (title: string, num: string, y: number, x0 = 72) => [
    ...mkText(title, x0, { y }).chars,
    ...mkText(num, 500, { y }).chars,
  ]

  it('detects ≥3 aligned rows with ascending page numbers (roman then arabic)', () => {
    const chars = [
      ...entryRow('ACKNOWLEDGEMENTS', 'iv', 700),
      ...entryRow('LIST OF TABLES', 'vi', 686),
      ...entryRow('CHAPTER 1. INTRODUCTION', '1', 672),
      ...entryRow('CHAPTER 2. RESULTS', '9', 658),
      ...mkText('Ordinary body paragraph follows here.', 72, { y: 600 }).chars,
    ]
    const { blocks, remainingUnits } = detectTocRows(unitsOf(chars))
    expect(blocks).toHaveLength(4)
    expect(blocks[0]!.tocEntry).toEqual({ level: 1, pageNumber: 'iv' })
    expect(blocks[2]!.lines[0]!.spans.map((s) => s.text).join('')).toBe('CHAPTER 1. INTRODUCTION')
    expect(remainingUnits.length).toBeGreaterThan(0) // the body line stays
  })

  it('rejects short runs and descending numbers', () => {
    const short = [...entryRow('Alpha', '3', 700), ...entryRow('Beta', '5', 686)]
    expect(detectTocRows(unitsOf(short)).blocks).toHaveLength(0)

    const descending = [
      ...entryRow('Alpha', '9', 700),
      ...entryRow('Beta', '5', 686),
      ...entryRow('Gamma', '3', 672),
    ]
    expect(detectTocRows(unitsOf(descending)).blocks).toHaveLength(0)
  })

  it('rejects rows whose numbers do not right-align', () => {
    const chars = [
      ...entryRow('Alpha', '3', 700),
      ...mkText('Beta', 72, { y: 686 }).chars,
      ...mkText('5', 300, { y: 686 }).chars,
      ...entryRow('Gamma', '7', 672),
    ]
    expect(detectTocRows(unitsOf(chars)).blocks).toHaveLength(0)
  })
})
