/** Merge-workbooks helpers: Excel-style name dedupe and IPC chunk sizing. */
import { describe, expect, it } from 'vitest'
import { chunkRowsFor, dedupeSheetName } from '../src/renderer/merge-workbooks'

describe('dedupeSheetName', () => {
  it('keeps an unused name', () => {
    expect(dedupeSheetName('Data', new Set(['Sheet1']))).toBe('Data')
  })

  it('suffixes like Excel and is case-insensitive', () => {
    expect(dedupeSheetName('Sheet1', new Set(['sheet1']))).toBe('Sheet1 (2)')
    expect(dedupeSheetName('Sheet1', new Set(['Sheet1', 'Sheet1 (2)']))).toBe('Sheet1 (3)')
  })

  it('trims the base so a suffixed name stays within Excel\u2019s 31-character cap', () => {
    const long = 'Regional Sales Summary FY2026 Q' // 31 chars, the maximum
    expect(long).toHaveLength(31)
    const deduped = dedupeSheetName(long, new Set([long]))
    expect(deduped).toBe('Regional Sales Summary FY20 (2)')
    expect(deduped).toHaveLength(31)
    // two-digit suffixes eat one more character of the base
    const taken = new Set([
      long,
      ...Array.from({ length: 8 }, (_, i) => `Regional Sales Summary FY20 (${i + 2})`),
    ])
    const tenth = dedupeSheetName(long, taken)
    expect(tenth).toBe('Regional Sales Summary FY2 (10)')
    expect(tenth).toHaveLength(31)
  })

  it('drops whitespace or an apostrophe exposed at the cut', () => {
    const spaceAtCut = 'Quarterly Budget Review 26 ABCD' // 31 chars, the cut lands after a space
    expect(spaceAtCut).toHaveLength(31)
    expect(dedupeSheetName(spaceAtCut, new Set([spaceAtCut]))).toBe(
      'Quarterly Budget Review 26 (2)',
    )
    const quoteAtCut = "Bob's Shop's Accounts FY26' XXX" // 31 chars, the cut lands after "'"
    expect(quoteAtCut).toHaveLength(31)
    expect(dedupeSheetName(quoteAtCut, new Set([quoteAtCut]))).toBe(
      "Bob's Shop's Accounts FY26 (2)",
    )
  })
})

describe('chunkRowsFor', () => {
  it('keeps rows × columns under the IPC cap', () => {
    expect(chunkRowsFor(1) * 1).toBeLessThanOrEqual(100_000)
    expect(chunkRowsFor(50) * 50).toBeLessThanOrEqual(100_000)
    expect(chunkRowsFor(200_000)).toBe(1)
  })
})
