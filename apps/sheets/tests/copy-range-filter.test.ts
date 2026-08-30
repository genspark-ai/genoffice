import { describe, expect, it, vi } from 'vitest'

import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'
import {
  expandToPrimitiveOps,
  filteredCopySourceRows,
  matchableCellText,
  workbookOperationSchema,
  type WorkbookOperation,
} from '../src/domain/workbook-dsl'
import { proposeOperations, type PlanContext } from '../src/renderer/plan-operations'

/// copy_range filterColumn/filterValues: row extraction for splitting data
/// by a column's values — matching rows land compacted at the target as
/// static values. The lazy executor shares filteredCopySourceRows /
/// matchableCellText with the demo path tested here; its chunked wiring is
/// exercised by the real-app driver.

describe('copy_range filter schema and geometry validation', () => {
  const base = {
    op: 'copy_range' as const,
    sheetId: 'sheet-1',
    source: 'A1:D6',
    target: 'F1',
  }

  it('parses filterColumn/filterValues', () => {
    const parsed = workbookOperationSchema.parse({
      ...base,
      filterColumn: 'D',
      filterValues: ['ja', 'ko'],
    })
    expect(parsed).toMatchObject({ filterColumn: 'D', filterValues: ['ja', 'ko'] })
  })

  it('rejects filterValues without filterColumn (and vice versa)', () => {
    expect(() =>
      expandToPrimitiveOps([{ ...base, filterValues: ['ja'] } as WorkbookOperation]),
    ).toThrow(/provided together/)
    expect(() =>
      expandToPrimitiveOps([{ ...base, filterColumn: 'D' } as WorkbookOperation]),
    ).toThrow(/provided together/)
  })

  it('rejects a filterColumn outside the source range', () => {
    expect(() =>
      expandToPrimitiveOps([
        { ...base, filterColumn: 'E', filterValues: ['ja'] } as WorkbookOperation,
      ]),
    ).toThrow(/outside the source range/)
  })

  it('rejects a multi-cell target on a filtered copy', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          ...base,
          target: 'F1:I6',
          filterColumn: 'D',
          filterValues: ['ja'],
        } as WorkbookOperation,
      ]),
    ).toThrow(/top-left cell/)
  })
})

describe('matchableCellText', () => {
  it('trims and lowercases text', () => {
    expect(matchableCellText('  JA ')).toBe('ja')
  })
  it('coerces numbers and booleans, and maps null to ""', () => {
    expect(matchableCellText(12.5)).toBe('12.5')
    expect(matchableCellText(true)).toBe('true')
    expect(matchableCellText(false)).toBe('false')
    expect(matchableCellText(null)).toBe('')
  })
})

describe('filteredCopySourceRows', () => {
  const values = new Map<string, string>([
    ['1:3', 'locale'],
    ['2:3', 'ja'],
    ['3:3', 'ko'],
    ['4:3', ' JA '],
    ['5:3', 'en'],
  ])
  const cellText = (row: number, column: number): string =>
    matchableCellText(values.get(`${row}:${column}`) ?? null)

  it('returns every source row when unfiltered', () => {
    const op = workbookOperationSchema.parse({
      op: 'copy_range',
      sheetId: 's1',
      source: 'A2:D4',
      target: 'F1',
    })
    expect(filteredCopySourceRows(op as never, cellText)).toEqual([1, 2, 3])
  })

  it('keeps only matching rows, trimmed and case-insensitive', () => {
    const op = workbookOperationSchema.parse({
      op: 'copy_range',
      sheetId: 's1',
      source: 'A1:D6',
      target: 'F1',
      filterColumn: 'D',
      filterValues: ['ja'],
    })
    // Screen rows 2 and 4 hold "ja" / " JA " (the header row does not match).
    expect(filteredCopySourceRows(op as never, cellText)).toEqual([2, 4])
  })

  it('returns null when nothing matches', () => {
    const op = workbookOperationSchema.parse({
      op: 'copy_range',
      sheetId: 's1',
      source: 'A1:D6',
      target: 'F1',
      filterColumn: 'D',
      filterValues: ['zh-CN'],
    })
    expect(filteredCopySourceRows(op as never, cellText)).toBeNull()
  })
})

describe('filtered copy_range on the demo workbook', () => {
  function demoContext(): PlanContext {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: {
            A1: { value: 'name' },
            B1: { value: 'locale' },
            A2: { value: 'ichiro' },
            B2: { value: 'ja' },
            A3: { value: 'minsu' },
            B3: { value: 'ko' },
            // Formula cell in a matching row: a filtered copy carries the
            // value, never the formula.
            A4: { value: 42, formula: '=6*7' },
            B4: { value: ' JA ' },
            A5: { value: 'david' },
            B5: { value: 'en' },
          },
        },
      ],
    })
    return {
      adapterRef: { current: adapter },
      univerRef: { current: null },
      lazyWorkbookRef: { current: null },
      lazyPreviewRef: { current: null },
      setPreview: vi.fn(),
      autoApplySafePlan: vi.fn().mockResolvedValue({ ok: true }),
    }
  }

  it('copies only matching rows, compacted, as static values', () => {
    const outcome = proposeOperations(
      demoContext(),
      [
        {
          op: 'copy_range',
          sheetId: 'sheet-1',
          source: 'A1:B5',
          target: 'D1',
          filterColumn: 'B',
          filterValues: ['ja'],
        },
      ],
      'split ja',
    )
    expect(outcome.ok).toBe(true)
    if (!outcome.ok) return
    const written = new Map(
      outcome.plan.cellChanges.map((change) => [change.address, change.after]),
    )
    // Rows 2 and 4 match and compact to target rows 1 and 2.
    expect(written.get('D1')).toEqual({ value: 'ichiro' })
    expect(written.get('E1')).toEqual({ value: 'ja' })
    expect(written.get('D2')).toEqual({ value: 42 })
    expect(written.get('D2')?.formula).toBeUndefined()
    expect(written.get('E2')).toEqual({ value: ' JA ' })
    // Non-matching rows (header, ko, en) do not copy.
    expect(written.size).toBe(4)
  })

  it('fails loud when the filter matches no rows', () => {
    const outcome = proposeOperations(
      demoContext(),
      [
        {
          op: 'copy_range',
          sheetId: 'sheet-1',
          source: 'A1:B5',
          target: 'D1',
          filterColumn: 'B',
          filterValues: ['zh-CN'],
        },
      ],
      'split zh',
    )
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('matched no source rows')
  })
})

describe('add_sheet rows/columns', () => {
  it('parses the optional grid size', () => {
    expect(
      workbookOperationSchema.parse({ op: 'add_sheet', name: 'ja', rows: 6700, columns: 8 }),
    ).toMatchObject({ rows: 6700, columns: 8 })
    expect(() => workbookOperationSchema.parse({ op: 'add_sheet', name: 'ja', rows: 0 })).toThrow()
    expect(() =>
      workbookOperationSchema.parse({ op: 'add_sheet', name: 'ja', rows: 2_000_000 }),
    ).toThrow()
  })

  it('records the requested grid size on the demo snapshot sheet', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [{ id: 'sheet-1', name: 'Sheet1', cells: { A1: { value: 'x' } } }],
    })
    const outcome = proposeOperations(
      {
        adapterRef: { current: adapter },
        univerRef: { current: null },
        lazyWorkbookRef: { current: null },
        lazyPreviewRef: { current: null },
        setPreview: vi.fn(),
        autoApplySafePlan: vi.fn().mockImplementation((plan) => {
          adapter.apply(plan)
          return Promise.resolve({ ok: true })
        }),
      },
      [{ op: 'add_sheet', name: 'ja', rows: 6700 }],
      'add sized sheet',
    )
    expect(outcome.ok).toBe(true)
    const added = adapter.getSnapshot().sheets.find((sheet) => sheet.name === 'ja')
    expect(added?.gridRows).toBe(6700)
  })
})
