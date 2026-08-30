import { describe, expect, it } from 'vitest'
import {
  decideLineShrinks,
  type ShrinkGap,
  type ShrinkLine,
} from '../src/renderer/editor/justify-shrink'

/**
 * Fixtures mirror the Word for Mac probe matrix (2026-08-27): TNR 14pt,
 * column 468pt, uniform 3.5pt spaces. A line of n words with the candidate
 * word overflowing the column by delta.
 */

const SP = 3.5
const COL = 468

function gap(width = SP, chars = 1, from = 0, to = 1): ShrinkGap {
  return { width, chars, from, to }
}

/** a justified line of `words` words whose pull candidate overflows by delta */
function line(words: number, delta: number, nextWordWidth: number, spaceW = SP): ShrinkLine {
  const gaps = Array.from({ length: words - 1 }, (_, i) => gap(spaceW, 1, i * 2, i * 2 + 1))
  const boundary = gap(spaceW, 1, 100, 101)
  // words sum so that: sum + interior + boundary + nextWord = COL + delta
  const sum = COL + delta - spaceW * words - nextWordWidth
  return {
    wordWidths: Array.from({ length: words }, () => sum / words),
    gaps,
    boundary,
    avail: COL,
    nextWordWidth,
  }
}

describe('decideLineShrinks — Word compat15 pull rule', () => {
  it('pulls within the tradeoff threshold (S=10, short word "de")', () => {
    // probe: w=13.21 pulls at delta=5.89, refuses at 6.27
    const [pull] = decideLineShrinks([line(10, 5.89, 13.21)])
    expect(pull).not.toBeNull()
    expect(pull!.gaps).toHaveLength(10)
    expect(pull!.perChar).toBeCloseTo((5.89 + 0.5) / 10, 5)
    const [no] = decideLineShrinks([line(10, 6.27, 13.21)])
    expect(no).toBeNull()
  })

  it('caps total shrink at 25% of the space width (S=10, "pour")', () => {
    // probe: w=25.66 pulls at 8.62 (cap 8.75), refuses at 9.00
    expect(decideLineShrinks([line(10, 8.62, 25.66)])[0]).not.toBeNull()
    expect(decideLineShrinks([line(10, 9.0, 25.66)])[0]).toBeNull()
  })

  it('cap scales with the space count (S=5 / S=15)', () => {
    // probe: S=5 pulls at 4.11 (cap 4.375), refuses at 4.44
    expect(decideLineShrinks([line(5, 4.11, 35)])[0]).not.toBeNull()
    expect(decideLineShrinks([line(5, 4.44, 35)])[0]).toBeNull()
    // probe: S=15 pulls at 12.89 (cap 13.125), refuses at 13.27
    expect(decideLineShrinks([line(15, 12.89, 35)])[0]).not.toBeNull()
    expect(decideLineShrinks([line(15, 13.27, 35)])[0]).toBeNull()
  })

  it('thresholds scale with the font size (28pt: space 7pt)', () => {
    // probe: 28pt "de" (26.43pt) pulls at 11.83, refuses at 12.59
    expect(decideLineShrinks([line(10, 11.83, 26.43, 7)])[0]).not.toBeNull()
    expect(decideLineShrinks([line(10, 12.59, 26.43, 7)])[0]).toBeNull()
  })

  it('keeps the compression of an already-pulled line', () => {
    const l = line(10, 5, 13.21)
    // simulate the pulled state: the word joined the line, natural > avail by 5
    l.wordWidths = [...l.wordWidths, 13.21]
    l.gaps = [...l.gaps, l.boundary!]
    l.boundary = null
    l.nextWordWidth = null
    const [keep] = decideLineShrinks([l])
    expect(keep).not.toBeNull()
    expect(keep!.perChar).toBeCloseTo((5 + 0.5) / 10, 3)
  })

  it('never pulls across a hard break or without a boundary space', () => {
    const l = line(10, 3, 13.21)
    l.boundary = null
    expect(decideLineShrinks([l])[0]).toBeNull()
  })

  it('leaves naturally fitting lines alone', () => {
    const l = line(10, -2, 13.21)
    expect(decideLineShrinks([l])[0]).toBeNull()
  })

  it('never divides by zero on a single-gap line', () => {
    const l: ShrinkLine = {
      wordWidths: [400],
      gaps: [],
      boundary: gap(),
      avail: COL,
      nextWordWidth: 60,
    }
    expect(decideLineShrinks([l])[0]).toBeNull()
  })
})
