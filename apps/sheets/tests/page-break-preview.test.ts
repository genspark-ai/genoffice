import { describe, expect, it } from 'vitest'

import {
  createEditJournal,
  recordPageSetup,
  recordStructuralOp,
} from '../src/renderer/edit-journal'
import {
  computePageBoundaries,
  effectivePageBreaks,
  printablePagePt,
} from '../src/renderer/page-break-preview'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

describe('printablePagePt', () => {
  it('defaults to A4 portrait with normal margins', () => {
    const page = printablePagePt({})
    expect(page.width).toBeCloseTo((8.27 - 1.4) * 72, 5)
    expect(page.height).toBeCloseTo((11.69 - 1.5) * 72, 5)
  })

  it('swaps the axes in landscape and honors paper/margin choices', () => {
    const page = printablePagePt({ paperSize: 1, orientation: 'landscape', margins: 'narrow' })
    expect(page.width).toBeCloseTo((11 - 0.5) * 72, 5)
    expect(page.height).toBeCloseTo((8.5 - 1.5) * 72, 5)
  })
})

describe('computePageBoundaries', () => {
  // 100px rows are 75pt scaled at 1; a 200pt page fits two rows.
  const rows100 = () => 100

  it('places automatic boundaries where the next row would overflow', () => {
    const boundaries = computePageBoundaries(rows100, 7, 200, 1, [])
    expect(boundaries).toEqual([
      { index: 2, manual: false },
      { index: 4, manual: false },
      { index: 6, manual: false },
    ])
  })

  it('manual breaks reset the accumulator and suppress the auto break there', () => {
    const boundaries = computePageBoundaries(rows100, 6, 200, 1, [1])
    expect(boundaries).toEqual([
      { index: 1, manual: true },
      { index: 3, manual: false },
      { index: 5, manual: false },
    ])
  })

  it('scale stretches page capacity and out-of-range manual ids are dropped', () => {
    expect(computePageBoundaries(rows100, 4, 200, 0.5, [0, 9])).toEqual([])
  })

  it('an oversized single row never yields a boundary at index 0', () => {
    expect(computePageBoundaries(() => 1000, 2, 200, 1, [])).toEqual([{ index: 1, manual: false }])
  })
})

function stateWith(setup: (state: LazyWorkbookState) => void): LazyWorkbookState {
  const state = {
    editJournal: createEditJournal(),
    sheetPageBreaks: new Map(),
  } as unknown as LazyWorkbookState
  setup(state)
  return state
}

describe('effectivePageBreaks', () => {
  it('is null before the sheet finishes indexing', () => {
    const state = stateWith(() => undefined)
    expect(effectivePageBreaks(state, 'sheet-1')).toBeNull()
  })

  it('prefers the journal set over the file set per axis', () => {
    const state = stateWith((current) => {
      current.sheetPageBreaks.set('sheet-1', { rowBreaks: [5], colBreaks: [2] })
      recordPageSetup(current.editJournal, 'sheet-1', { rowBreaks: [9] })
    })
    expect(effectivePageBreaks(state, 'sheet-1')).toEqual({ rowBreaks: [9], colBreaks: [2] })
  })

  it('shifts file breaks through structural ops and drops deleted ones', () => {
    const state = stateWith((current) => {
      current.sheetPageBreaks.set('sheet-1', { rowBreaks: [5, 10], colBreaks: [] })
      recordStructuralOp(
        current.editJournal,
        'sheet-1',
        { kind: 'remove-rows', index: 4, count: 3 },
        'Sheet1',
      )
    })
    // Row 5 fell inside the deleted band; row 10 shifts up by 3.
    expect(effectivePageBreaks(state, 'sheet-1')).toEqual({ rowBreaks: [7], colBreaks: [] })
  })
})
