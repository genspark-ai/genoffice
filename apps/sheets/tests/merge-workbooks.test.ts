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
})

describe('chunkRowsFor', () => {
  it('keeps rows × columns under the IPC cap', () => {
    expect(chunkRowsFor(1) * 1).toBeLessThanOrEqual(100_000)
    expect(chunkRowsFor(50) * 50).toBeLessThanOrEqual(100_000)
    expect(chunkRowsFor(200_000)).toBe(1)
  })
})
