import { describe, expect, it } from 'vitest'

import { computeSortChanges } from '../src/domain/sort-range'
import type { CellState } from '../src/domain/workbook.types'

function reader(cells: Record<string, CellState>) {
  return (address: string): CellState => cells[address] ?? { value: null }
}

const DATA: Record<string, CellState> = {
  A1: { value: 'Name' },
  B1: { value: 'Score' },
  A2: { value: 'Cara' },
  B2: { value: 70 },
  A3: { value: 'Ann' },
  B3: { value: 90 },
  A4: { value: 'Bob' },
  B4: { value: 80 },
}

describe('computeSortChanges', () => {
  it('sorts rows descending by the key column, keeping rows intact', () => {
    const changes = computeSortChanges(
      { range: 'A1:B4', byColumn: 'B', ascending: false, hasHeader: true },
      reader(DATA),
    )
    const byAddress = Object.fromEntries(changes.map((change) => [change.address, change.after]))
    expect(byAddress).toEqual({
      A2: 'Ann',
      B2: 90,
      A3: 'Bob',
      B3: 80,
      A4: 'Cara',
      B4: 70,
    })
    expect(changes.every((change) => change.before !== change.after)).toBe(true)
  })

  it('emits nothing when the data is already in order', () => {
    const ordered: Record<string, CellState> = {
      A1: { value: 1 },
      A2: { value: 2 },
      A3: { value: 3 },
    }
    const changes = computeSortChanges(
      { range: 'A1:A3', byColumn: 'A', ascending: true, hasHeader: false },
      reader(ordered),
    )
    expect(changes).toEqual([])
  })

  it('keeps blanks last regardless of direction (only changed cells emitted)', () => {
    const cells: Record<string, CellState> = {
      A1: { value: null },
      A2: { value: 5 },
      A3: { value: 1 },
    }
    const asc = computeSortChanges(
      { range: 'A1:A3', byColumn: 'A', ascending: true, hasHeader: false },
      reader(cells),
    )
    expect(Object.fromEntries(asc.map((c) => [c.address, c.after]))).toEqual({
      A1: 1,
      A3: null,
    })
    const desc = computeSortChanges(
      { range: 'A1:A3', byColumn: 'A', ascending: false, hasHeader: false },
      reader(cells),
    )
    expect(Object.fromEntries(desc.map((c) => [c.address, c.after]))).toEqual({
      A1: 5,
      A2: 1,
      A3: null,
    })
  })

  it('rejects ranges containing formulas', () => {
    const cells: Record<string, CellState> = {
      A1: { value: 2 },
      A2: { value: null, formula: '=B1*2' },
    }
    expect(() =>
      computeSortChanges(
        { range: 'A1:A2', byColumn: 'A', ascending: true, hasHeader: false },
        reader(cells),
      ),
    ).toThrow(/formula/i)
  })

  it('rejects a key column outside the range and header-only ranges', () => {
    expect(() =>
      computeSortChanges(
        { range: 'A1:B3', byColumn: 'C', ascending: true, hasHeader: false },
        reader(DATA),
      ),
    ).toThrow(/outside/)
    expect(() =>
      computeSortChanges(
        { range: 'A1:B2', byColumn: 'A', ascending: true, hasHeader: true },
        reader(DATA),
      ),
    ).toThrow(/two data rows/)
  })

  it('sorts numbers before text and uses numeric-aware string order', () => {
    const cells: Record<string, CellState> = {
      A1: { value: 'item10' },
      A2: { value: 3 },
      A3: { value: 'item2' },
    }
    const changes = computeSortChanges(
      { range: 'A1:A3', byColumn: 'A', ascending: true, hasHeader: false },
      reader(cells),
    )
    expect(Object.fromEntries(changes.map((c) => [c.address, c.after]))).toEqual({
      A1: 3,
      A2: 'item2',
      A3: 'item10',
    })
  })
})
