import { beforeEach, describe, expect, it } from 'vitest'

import {
  applyRowProperties,
  measureWrapAutoFitRows,
  numericWrapOverride,
  resetStaleWrapAutoHeights,
  takeContaminatedRows,
  trackPreIndexMeasuredRows,
  wrapAutoFitRows,
  wrapMeasureGate,
} from '../src/renderer/univer-sync'
import { journalSuppression, loadAutoHeightSuppression } from '../src/renderer/univer-state'

function makeWorksheet() {
  const calls: Array<{ method: string; row: number; px?: number; suppressed?: boolean }> = []
  const worksheet = {
    setRowHeightsForced: (row: number, _count: number, px: number) =>
      calls.push({ method: 'forced', row, px }),
    // Re-enables ia only; the load gate must be up so the command cannot
    // measure (and balloon) the row while file content installs.
    setRowAutoHeight: (row: number) =>
      calls.push({ method: 'auto', row, suppressed: loadAutoHeightSuppression.active }),
    hideRows: (row: number) => calls.push({ method: 'hide', row }),
    getSheet: () => ({ setRowStyle: () => undefined }),
  }
  return { worksheet, calls }
}

function makeState(defaultRowHeight: number | null = null) {
  return {
    file: { styles: [], sheets: [{ id: 'sheet-1', defaultRowHeight }] },
    appliedRowKeys: new Map<string, Set<string>>(),
    hiddenFileRows: new Map<string, Set<number>>(),
    outline: new Map(),
  }
}

describe('applyRowProperties', () => {
  it('applies stored heights verbatim and never re-measures on open', () => {
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState() as never, 'sheet-1', [
      // customHeight="1": the user fixed it — clip like Excel, stay locked.
      { row: 0, height: 56, customHeight: true, hidden: false },
      // Plain ht: Excel renders the stored value as-is on open; the row only
      // goes back to auto mode so a LATER user edit can re-fit it.
      { row: 1, height: 38.25, hidden: false },
      { row: 2, hidden: true },
    ] as never)
    expect(calls).toEqual([
      { method: 'forced', row: 0, px: 75 }, // 56pt → 75px
      { method: 'forced', row: 1, px: 51 }, // 38.25pt → 51px
      { method: 'auto', row: 1, suppressed: true },
      { method: 'hide', row: 2 },
    ])
  })

  it('records file-hidden rows so the viewport loader can budget by visible rows', () => {
    const { worksheet } = makeWorksheet()
    const state = makeState()
    applyRowProperties(worksheet as never, state as never, 'sheet-1', [
      { row: 0, hidden: false },
      { row: 1, hidden: true },
      { row: 2, hidden: true },
    ] as never)
    expect([...(state.hiddenFileRows.get('sheet-1') ?? [])]).toEqual([1, 2])
  })

  it('drops the suppression flag after the rows are applied', () => {
    const { worksheet } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState() as never, 'sheet-1', [
      { row: 0, height: 20, hidden: false },
    ] as never)
    expect(loadAutoHeightSuppression.active).toBe(false)
  })

  it('keeps sub-default heights locked — spacer rows are not auto-fit results', () => {
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState() as never, 'sheet-1', [
      // Print-style layouts build vertical rhythm from tiny rows; an edit-time
      // auto-fit would balloon each to a full text line.
      { row: 0, height: 2.25, hidden: false },
      { row: 1, height: 0.95, hidden: false },
      // At/above one line: back to auto mode for future edits.
      { row: 2, height: 18, hidden: false },
    ] as never)
    expect(calls).toEqual([
      { method: 'forced', row: 0, px: 3 },
      { method: 'forced', row: 1, px: 1 },
      { method: 'forced', row: 2, px: 24 },
      { method: 'auto', row: 2, suppressed: true },
    ])
  })

  it('treats Excel-default rows as auto rows when the file omits the default', () => {
    // No sheetFormatPr default → the cutoff is Excel's factory 15pt (20px),
    // not Univer's taller UI default, so ordinary 15pt rows keep auto mode.
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState(null) as never, 'sheet-1', [
      { row: 0, height: 15, hidden: false },
    ] as never)
    expect(calls).toEqual([
      { method: 'forced', row: 0, px: 20 },
      { method: 'auto', row: 0, suppressed: true },
    ])
  })

  it('compares a row at a fractional default as at-default, not below', () => {
    // 14.3pt → 19.07px: both sides round to 19, so the row is not treated as
    // a spacer.
    const { worksheet, calls } = makeWorksheet()
    applyRowProperties(worksheet as never, makeState(14.3) as never, 'sheet-1', [
      { row: 0, height: 14.3, hidden: false },
    ] as never)
    expect(calls).toEqual([
      { method: 'forced', row: 0, px: 19 },
      { method: 'auto', row: 0, suppressed: true },
    ])
  })

  it('re-applies when the customHeight flag changes but dedupes repeats', () => {
    const { worksheet, calls } = makeWorksheet()
    const state = makeState()
    const rows = [{ row: 0, height: 56, hidden: false }] as never
    applyRowProperties(worksheet as never, state as never, 'sheet-1', rows)
    applyRowProperties(worksheet as never, state as never, 'sheet-1', rows)
    expect(calls).toHaveLength(2)
    applyRowProperties(worksheet as never, state as never, 'sheet-1', [
      { row: 0, height: 56, customHeight: true, hidden: false },
    ] as never)
    expect(calls).toEqual([
      { method: 'forced', row: 0, px: 75 },
      { method: 'auto', row: 0, suppressed: true },
      { method: 'forced', row: 0, px: 75 },
    ])
  })
})

describe('numericWrapOverride', () => {
  it('unwraps numeric cells with a wrap style — Excel never wraps numbers', () => {
    expect(numericWrapOverride(45_123.87, true)).toBe(true)
  })

  it('leaves text, booleans, and non-wrap styles alone', () => {
    expect(numericWrapOverride('long text', true)).toBe(false)
    expect(numericWrapOverride(true, true)).toBe(false)
    expect(numericWrapOverride(null, true)).toBe(false)
    expect(numericWrapOverride(45_123.87, false)).toBe(false)
    expect(numericWrapOverride(45_123.87, undefined)).toBe(false)
  })
})

describe('wrapAutoFitRows', () => {
  const styles = [{ wrapText: true }, { wrapText: false }, {}] as never as Parameters<
    typeof wrapAutoFitRows
  >[1]
  const range = { startRow: 0, endRow: 99, startColumn: 0, endColumn: 9 }

  it('selects auto rows whose wrap cells hold text (prod_026/prod_006 shape)', () => {
    const rows = wrapAutoFitRows(
      [
        { row: 0, column: 0, value: 'wrapping header text', styleIndex: 0 },
        { row: 1, column: 0, value: 'more wrapping text', styleIndex: 0 },
        { row: 2, column: 0, value: 'no wrap style', styleIndex: 1 },
        { row: 3, column: 0, value: 'style-less', styleIndex: 2 },
        { row: 4, column: 0, value: 'no style at all' },
      ] as never,
      styles,
      [],
      [],
      false,
      null,
      range,
    )
    expect(rows).toEqual([0, 1])
  })

  it('re-fits cached-ht auto rows but keeps user-fixed and spacer rows (prod_054)', () => {
    const rows = wrapAutoFitRows(
      [
        { row: 0, column: 0, value: 'wrapping text', styleIndex: 0 },
        { row: 1, column: 0, value: 'wrapping text', styleIndex: 0 },
        { row: 2, column: 0, value: 'wrapping text', styleIndex: 0 },
      ] as never,
      styles,
      // A cached ht without customHeight stays in auto mode — Excel re-fits
      // it on open (live probe: ht="30" rows reopen at 16pt). customHeight
      // and sub-default spacer heights stay verbatim.
      [
        { row: 0, height: 30, hidden: false },
        { row: 1, height: 56, customHeight: true, hidden: false },
        { row: 2, height: 6, hidden: false },
      ] as never,
      [],
      false,
      15,
      range,
    )
    expect(rows).toEqual([0])
  })

  it('selects nothing when sheetFormatPr customHeight fixes the default (prod_012)', () => {
    const rows = wrapAutoFitRows(
      [{ row: 0, column: 0, value: 'wrapping text', styleIndex: 0 }] as never,
      styles,
      [],
      [],
      true,
      null,
      range,
    )
    expect(rows).toEqual([])
  })

  it('skips numeric wrap cells — Excel never wraps numbers (prod_098 ####)', () => {
    const rows = wrapAutoFitRows(
      [
        { row: 0, column: 0, value: 46_234.86, styleIndex: 0 },
        { row: 1, column: 0, value: '', styleIndex: 0 },
        { row: 2, column: 0, styleIndex: 0 },
      ] as never,
      styles,
      [],
      [],
      false,
      null,
      range,
    )
    expect(rows).toEqual([])
  })

  it('ignores cells outside the patched range', () => {
    const rows = wrapAutoFitRows(
      [
        { row: 100, column: 0, value: 'wrapping text', styleIndex: 0 },
        { row: 5, column: 20, value: 'wrapping text', styleIndex: 0 },
      ] as never,
      styles,
      [],
      [],
      false,
      null,
      range,
    )
    expect(rows).toEqual([])
  })

  it('honors wrap inherited from customFormat row and column styles', () => {
    const rows = wrapAutoFitRows(
      [
        { row: 0, column: 0, value: 'row-style wrap' },
        { row: 1, column: 3, value: 'col-style wrap' },
        { row: 2, column: 3, value: 'own no-wrap xf wins', styleIndex: 1 },
        { row: 3, column: 3, value: 'no-wrap row xf blocks col wrap' },
      ] as never,
      styles,
      [
        { row: 0, hidden: false, styleIndex: 0 },
        { row: 3, hidden: false, styleIndex: 1 },
      ] as never,
      [{ startColumn: 2, endColumn: 5, hidden: false, styleIndex: 0 }] as never,
      false,
      null,
      range,
    )
    expect(rows).toEqual([0, 1])
  })

  it('lets a later overlapping no-wrap column span override an earlier wrap span', () => {
    const rows = wrapAutoFitRows(
      [{ row: 0, column: 3, value: 'later span painted no-wrap' }] as never,
      styles,
      [],
      [
        { startColumn: 2, endColumn: 5, hidden: false, styleIndex: 0 },
        { startColumn: 3, endColumn: 4, hidden: false, styleIndex: 1 },
      ] as never,
      false,
      null,
      range,
    )
    expect(rows).toEqual([])
  })
})

describe('measureWrapAutoFitRows', () => {
  beforeEach(() => {
    wrapMeasureGate.ready = true
    wrapMeasureGate.pending.length = 0
  })

  it('queues measures until the auto-height interceptor exists (Rendered)', () => {
    wrapMeasureGate.ready = false
    const worksheet = {
      setRowAutoHeight: () => {
        throw new Error('must not measure before lifecycle Rendered')
      },
    }
    measureWrapAutoFitRows(worksheet as never, [1, 2])
    expect(wrapMeasureGate.pending).toHaveLength(1)
    expect(wrapMeasureGate.pending[0]?.rows).toEqual([1, 2])
  })

  it('batches contiguous rows and measures through the user-autofit channel', () => {
    const calls: Array<{ start: number; count: number; journal: boolean; gate: boolean }> = []
    const worksheet = {
      setRowAutoHeight: (start: number, count: number) =>
        calls.push({
          start,
          count,
          // Undo/journal must stay quiet while opening a file...
          journal: journalSuppression.active,
          // ...but the load gate must be DOWN, or the measure yields nothing.
          gate: loadAutoHeightSuppression.active,
        }),
    }
    measureWrapAutoFitRows(worksheet as never, [2, 3, 4, 7, 9, 10])
    expect(calls).toEqual([
      { start: 2, count: 3, journal: true, gate: false },
      { start: 7, count: 1, journal: true, gate: false },
      { start: 9, count: 2, journal: true, gate: false },
    ])
    expect(journalSuppression.active).toBe(false)
  })

  it('does nothing for an empty row set', () => {
    const worksheet = {
      setRowAutoHeight: () => {
        throw new Error('must not measure')
      },
    }
    measureWrapAutoFitRows(worksheet as never, [])
  })
})

describe('merged wrap cells and stale auto heights (prod_100 shape)', () => {
  const styles = [{ wrapText: true }] as never as Parameters<typeof wrapAutoFitRows>[1]
  const range = { startRow: 0, endRow: 9, startColumn: 0, endColumn: 19 }

  it('merge-covered wrap cells do not qualify their row', () => {
    const cells = [
      { row: 0, column: 8, value: 'أسم المنشأة:', styleIndex: 0 },
      { row: 1, column: 8, value: 'رقم إشتراك المنشأة:', styleIndex: 0 },
      { row: 3, column: 0, value: 'unmerged wrapping text', styleIndex: 0 },
    ] as never
    const merges = [
      { startRow: 0, endRow: 0, startColumn: 8, endColumn: 13 },
      { startRow: 1, endRow: 1, startColumn: 8, endColumn: 13 },
    ]
    expect(wrapAutoFitRows(cells, styles, [], [], false, null, range, merges)).toEqual([3])
    expect(wrapAutoFitRows(cells, styles, [], [], false, null, range)).toEqual([0, 1, 3])
  })

  it('resetStaleWrapAutoHeights clears tracked contaminated rows only', () => {
    const executed: unknown[] = []
    const runtime = {
      univerAPI: {
        syncExecuteCommand: (id: string, params: unknown) => executed.push([id, params]),
      },
    } as never
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getSheet: () => ({
        getSnapshot: () => ({
          rowData: { 0: { ah: 194 }, 1: { ah: 306 }, 5: { ah: 22 }, 6: { ah: 250 }, 7: { ah: 8 } },
        }),
      }),
    } as never
    // Queued pre-Rendered measure must lose the reset rows or the lifecycle
    // flush re-poisons them (bugbot).
    wrapMeasureGate.pending.push({ worksheet, rows: [0, 1, 3] })
    resetStaleWrapAutoHeights(
      runtime,
      'file-x',
      worksheet,
      [0, 1, 5, 6, 7],
      [
        { row: 5, height: 22, customHeight: true },
        // Cached ht without customHeight stays auto mode: a poisoned merged
        // measure on such a row must reset too (bugbot).
        { row: 6, height: 30, customHeight: false },
        // Sub-default spacer rows keep their stored height verbatim.
        { row: 7, height: 8, customHeight: false },
      ] as never,
      15,
    )
    expect(executed).toEqual([
      [
        'sheet.mutation.set-worksheet-row-auto-height',
        {
          unitId: 'file-x',
          subUnitId: 'sheet-1',
          rowsAutoHeightInfo: [
            { row: 0, autoHeight: undefined },
            { row: 1, autoHeight: undefined },
            { row: 6, autoHeight: undefined },
          ],
        },
      ],
    ])
    expect(wrapMeasureGate.pending.pop()?.rows).toEqual([3])
  })

  it('takeContaminatedRows returns tracked rows that lost qualification to merges', () => {
    trackPreIndexMeasuredRows('k:s', [0, 1, 4])
    // Row 0 still qualifies (kept, untracked); row 1 qualified only without
    // merges (contaminated); row 4's cells are outside this window (stays
    // tracked); row 9 was never measured pre-index (ignored).
    expect(takeContaminatedRows('k:s', [0, 1, 9], [0, 9])).toEqual([1])
    expect(takeContaminatedRows('k:s', [0, 1, 9], [0, 9])).toEqual([])
    expect(takeContaminatedRows('k:s', [4], [])).toEqual([4])
  })
})
