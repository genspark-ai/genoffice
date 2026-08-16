import { describe, expect, it } from 'vitest'

import { applyRowProperties } from '../src/renderer/univer-sync'

function makeWorksheet() {
  const calls: Array<{ method: string; row: number; px?: number }> = []
  const worksheet = {
    setRowHeightsForced: (row: number, _count: number, px: number) =>
      calls.push({ method: 'forced', row, px }),
    setRowHeights: (row: number, _count: number, px: number) =>
      calls.push({ method: 'auto-floor', row, px }),
    hideRows: (row: number) => calls.push({ method: 'hide', row }),
    getSheet: () => ({ setRowStyle: () => undefined }),
  }
  return { worksheet, calls }
}

function makeState(defaultRowHeight: number | null = null) {
  return {
    file: { styles: [], sheets: [{ id: 'sheet-1', defaultRowHeight }] },
    appliedRowKeys: new Map<string, Set<string>>(),
    outline: new Map(),
  }
}

describe('applyRowProperties', () => {
  it('forces customHeight rows and keeps auto-height for auto-fit ht rows', () => {
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState() as never, 'sheet-1', [
      // customHeight="1": the user fixed it — clip like Excel.
      { row: 0, height: 56, customHeight: true, hidden: false },
      // Plain ht: Excel's last auto-fit; content may still grow the row.
      { row: 1, height: 38.25, hidden: false },
      { row: 2, hidden: true },
    ] as never)
    expect(calls).toEqual([
      { method: 'forced', row: 0, px: 75 }, // 56pt → 75px
      { method: 'auto-floor', row: 1, px: 51 }, // 38.25pt → 51px
      { method: 'hide', row: 2 },
    ])
  })

  it('forces sub-default heights — spacer rows are not auto-fit results', () => {
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState() as never, 'sheet-1', [
      // Print-style layouts build vertical rhythm from tiny rows; auto-height
      // would balloon each to a full text line.
      { row: 0, height: 2.25, hidden: false },
      { row: 1, height: 0.95, hidden: false },
      // At/above one line: still auto-growable for wrapped content.
      { row: 2, height: 18, hidden: false },
    ] as never)
    expect(calls).toEqual([
      { method: 'forced', row: 0, px: 3 },
      { method: 'forced', row: 1, px: 1 },
      { method: 'auto-floor', row: 2, px: 24 },
    ])
  })

  it('treats Excel-default rows as auto-growable when the file omits the default', () => {
    // No sheetFormatPr default → the cutoff is Excel's factory 15pt (20px),
    // not Univer's taller UI default, so ordinary 15pt rows keep auto-grow.
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState(null) as never, 'sheet-1', [
      { row: 0, height: 15, hidden: false },
    ] as never)
    expect(calls).toEqual([{ method: 'auto-floor', row: 0, px: 20 }])
  })

  it('compares a row at a fractional default as at-default, not below', () => {
    // 14.3pt → 19.07px: both sides round to 19, so the row is not forced.
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState(14.3) as never, 'sheet-1', [
      { row: 0, height: 14.3, hidden: false },
    ] as never)
    expect(calls).toEqual([{ method: 'auto-floor', row: 0, px: 19 }])
  })

  it('re-applies when the customHeight flag changes but dedupes repeats', () => {
    const { worksheet, calls } = makeWorksheet()
    const state = makeState()
    const rows = [{ row: 0, height: 56, hidden: false }] as never
    applyRowProperties(worksheet as never, state as never, 'sheet-1', rows)
    applyRowProperties(worksheet as never, state as never, 'sheet-1', rows)
    expect(calls).toHaveLength(1)
    applyRowProperties(worksheet as never, state as never, 'sheet-1', [
      { row: 0, height: 56, customHeight: true, hidden: false },
    ] as never)
    expect(calls).toEqual([
      { method: 'auto-floor', row: 0, px: 75 },
      { method: 'forced', row: 0, px: 75 },
    ])
  })
})
