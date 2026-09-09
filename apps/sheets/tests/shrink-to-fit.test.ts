import { describe, expect, it } from 'vitest'

import { shrinkToFitApplies, shrinkToFitFontSize } from '../src/renderer/univer-sync'
import type { WorkbookCellStyle } from '../src/shared/desktop-api'

// 10px per character at the base size; scales linearly with font size.
const measure = (line: string): number => line.length * 10

describe('shrinkToFitFontSize', () => {
  it('returns null when the text already fits', () => {
    expect(shrinkToFitFontSize('1980', 12, 50, measure)).toBeNull()
  })

  it('scales the font down proportionally and floors to an integer', () => {
    // 5 chars * 10px = 50px into 42px: 12 * 42/50 = 10.08 -> 10
    expect(shrinkToFitFontSize('1980X', 12, 42, measure)).toBe(10)
  })

  it('uses the widest line of a multiline cell', () => {
    expect(shrinkToFitFontSize('ab\r\nabcdef', 12, 30, measure)).toBe(6)
  })

  it('clamps at 1pt and rejects a zero-width column', () => {
    expect(shrinkToFitFontSize('abcdefghij', 12, 1, measure)).toBe(1)
    expect(shrinkToFitFontSize('abc', 12, 0, measure)).toBeNull()
  })
})

describe('shrinkToFitApplies', () => {
  const base: WorkbookCellStyle = {
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    wrapText: false,
    diagonalUp: false,
    diagonalDown: false,
  }
  const shrink: WorkbookCellStyle = { ...base, shrinkToFit: true }

  it('applies to plain shrinkToFit text', () => {
    expect(shrinkToFitApplies(shrink, 'MOBILIZATION', false)).toBe(true)
  })

  it('yields to wrapText when both flags are set (Excel greys shrink out)', () => {
    expect(shrinkToFitApplies({ ...shrink, wrapText: true }, 'MOBILIZATION', false)).toBe(false)
  })

  it('skips empty text, pending formulas and unstyled cells', () => {
    expect(shrinkToFitApplies(shrink, '', false)).toBe(false)
    expect(shrinkToFitApplies(shrink, '=A1', true)).toBe(false)
    expect(shrinkToFitApplies(undefined, 'x', false)).toBe(false)
    expect(shrinkToFitApplies({ ...base, wrapText: true }, 'x', false)).toBe(false)
  })
})
