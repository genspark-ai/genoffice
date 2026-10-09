import { describe, expect, it } from 'vitest'

import {
  FORMULA_MODE_MAX_CELLS,
  FULL_LOAD_MAX_CELLS,
  FULL_LOAD_MAX_GRID_CELLS,
  FULL_LOAD_MAX_STORED_CELLS,
} from '../src/renderer/app-constants'
import { fitsFullLoad, opensInFormulaMode, workbookCellCounts } from '../src/renderer/load-budget'

describe('workbookCellCounts', () => {
  it('sums the bounding box and the stored count separately', () => {
    const counts = workbookCellCounts([
      { rowCount: 50_000, columnCount: 5_600, storedCellCount: 249_000 },
      { rowCount: 10, columnCount: 10, storedCellCount: 7 },
    ])
    expect(counts.gridCells).toBe(280_000_100)
    expect(counts.storedCells).toBe(249_007)
  })

  it('falls back to the bounding box for sheets without a reported count', () => {
    const counts = workbookCellCounts([
      { rowCount: 100, columnCount: 10 },
      { rowCount: 10, columnCount: 10, storedCellCount: 0 },
    ])
    expect(counts.storedCells).toBe(1_000)
    expect(counts.valueCells).toBe(1_000)
  })

  it('uses the value extent for the box and the value count when reported', () => {
    const counts = workbookCellCounts([
      {
        rowCount: 1_048_000,
        columnCount: 27,
        storedCellCount: 567_001,
        valueCellCount: 567_000,
        valueRowCount: 21_000,
        valueColumnCount: 27,
      },
      { rowCount: 10, columnCount: 10, storedCellCount: 7 },
    ])
    expect(counts.gridCells).toBe(21_000 * 27 + 100)
    expect(counts.storedCells).toBe(567_008)
    expect(counts.valueCells).toBe(567_007)
  })
})

describe('load decisions', () => {
  it('a sparse sheet fully loads while its box stays within the dense-install ceiling', () => {
    const counts = workbookCellCounts([
      { rowCount: 25_000, columnCount: 100, storedCellCount: 249_000 },
    ])
    expect(counts.gridCells).toBeGreaterThan(FULL_LOAD_MAX_CELLS / 4)
    expect(fitsFullLoad(counts)).toBe(true)
    expect(opensInFormulaMode(counts)).toBe(false)
  })

  it('a box past the ceiling is too large however few cells it stores', () => {
    const counts = workbookCellCounts([
      { rowCount: 50_000, columnCount: 5_600, storedCellCount: 249_000 },
    ])
    expect(counts.gridCells).toBeGreaterThan(FULL_LOAD_MAX_GRID_CELLS)
    expect(fitsFullLoad(counts)).toBe(false)
    expect(
      opensInFormulaMode(
        workbookCellCounts([{ rowCount: 50_000, columnCount: 5_600, storedCellCount: 10 }]),
      ),
    ).toBe(false)
  })

  it('a few stored cells in a wide box open in formula mode', () => {
    const counts = workbookCellCounts([
      { rowCount: 1_000, columnCount: 1_000, storedCellCount: FORMULA_MODE_MAX_CELLS },
    ])
    expect(opensInFormulaMode(counts)).toBe(true)
    expect(
      opensInFormulaMode(
        workbookCellCounts([
          { rowCount: 1_000, columnCount: 1_000, storedCellCount: FORMULA_MODE_MAX_CELLS + 1 },
        ]),
      ),
    ).toBe(false)
  })

  it('a dense sheet past the cap is too large even with a small box', () => {
    const counts = workbookCellCounts([
      { rowCount: 2_000, columnCount: 600, storedCellCount: FULL_LOAD_MAX_CELLS + 1 },
    ])
    expect(fitsFullLoad(counts)).toBe(false)
  })

  it('style-only blanks below the data do not deny the offer (genoffice#1927)', () => {
    const counts = workbookCellCounts([
      {
        rowCount: 61_000,
        columnCount: 27,
        storedCellCount: 1_647_000,
        valueCellCount: 567_000,
        valueRowCount: 21_000,
        valueColumnCount: 27,
      },
    ])
    expect(counts.storedCells).toBeGreaterThan(FULL_LOAD_MAX_CELLS)
    expect(fitsFullLoad(counts)).toBe(true)
  })

  it('a stray formatted cell far below the data does not deny the offer', () => {
    const counts = workbookCellCounts([
      {
        rowCount: 1_048_000,
        columnCount: 27,
        storedCellCount: 567_001,
        valueCellCount: 567_000,
        valueRowCount: 21_000,
        valueColumnCount: 27,
      },
    ])
    expect(1_048_000 * 27).toBeGreaterThan(FULL_LOAD_MAX_GRID_CELLS)
    expect(fitsFullLoad(counts)).toBe(true)
  })

  it('millions of style-only blanks still exceed the stored ceiling', () => {
    const counts = workbookCellCounts([
      {
        rowCount: 987_813,
        columnCount: 18,
        storedCellCount: FULL_LOAD_MAX_STORED_CELLS + 1,
        valueCellCount: 3_596,
        valueRowCount: 430,
        valueColumnCount: 18,
      },
    ])
    expect(fitsFullLoad(counts)).toBe(false)
  })

  it('totals across sheets decide, not the largest sheet', () => {
    const half = Math.ceil(FULL_LOAD_MAX_CELLS / 2) + 1
    const counts = workbookCellCounts([
      { rowCount: 1, columnCount: 1, storedCellCount: half },
      { rowCount: 1, columnCount: 1, storedCellCount: half },
    ])
    expect(fitsFullLoad(counts)).toBe(false)
  })
})
