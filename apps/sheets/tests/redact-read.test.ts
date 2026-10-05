import { describe, expect, it, vi } from 'vitest'

import { createRangeAggregator, formatRangeAggregate } from '../src/renderer/ai/aggregate'
import {
  NO_REDACTIONS,
  buildRedactionIndex,
  countWithheldIn,
  placeholderInstruction,
  placeholderSource,
  sanitizeLabel,
} from '../src/renderer/ai/redact'
import {
  readCells,
  redactionsOf,
  type WorkbookReadContext,
} from '../src/renderer/ai/workbook-readers'
import type { SheetRedactionState } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'

const SHEETS = [
  { id: 'sh1', name: 'Customers' },
  { id: 'sh2', name: 'Orders' },
]

const STATES: SheetRedactionState[] = [
  {
    sheetName: 'Customers',
    // B2 only — the phone number the reader does not want sent anywhere.
    marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: '客户电话' }],
  },
]

const index = buildRedactionIndex(STATES, SHEETS)

/// A demo-mode context: the in-memory adapter holds every cell, which is the
/// path the projection has to hold without a live Univer grid.
function demoContext(cells: Record<string, unknown>, redactions = index): WorkbookReadContext {
  const worksheet = {
    getSheetId: () => 'sh1',
    getSheetName: () => 'Customers',
    getRange: () => ({ getValue: () => null }),
  }
  return {
    univerRef: {
      current: {
        univerAPI: {
          getActiveWorkbook: () => ({
            getActiveSheet: () => worksheet,
            getSheetBySheetId: () => worksheet,
            getActiveRange: () => null,
            getSheets: () => [worksheet],
          }),
        },
      },
    } as never,
    lazyWorkbookRef: { current: null },
    adapterRef: {
      current: {
        getSnapshot: () => ({
          revision: 1,
          sheets: [{ id: 'sh1', name: 'Customers', cells }],
        }),
      },
    } as never,
    redactionIndexRef: { current: redactions },
  }
}

describe('the withheld-cell index', () => {
  it('answers by sheet id, which is not the name the part is keyed by', () => {
    // The part is keyed by name; every read tool addresses sheets by id. If the
    // translation were missing, the mark would never match a cell.
    expect(index.labelAt('sh1', 1, 1)).toBe('客户电话')
    expect(index.labelAt('sh2', 1, 1)).toBeNull()
  })

  it('leaves the cells around a mark alone', () => {
    expect(index.labelAt('sh1', 0, 1)).toBeNull()
    expect(index.labelAt('sh1', 1, 0)).toBeNull()
    expect(index.labelAt('sh1', 1, 2)).toBeNull()
    expect(index.labelAt('sh1', 2, 1)).toBeNull()
  })

  it('drops a mark naming a sheet this workbook does not have', () => {
    // Applying one sheet's marks by position would withhold the wrong cells —
    // wrong in the safe direction for some cells and unsafe for others. There
    // is no position to fall back on, so the mark is not applied at all.
    const orphaned = buildRedactionIndex(
      [
        {
          sheetName: 'Deleted',
          marks: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0, label: 'x' }],
        },
      ],
      SHEETS,
    )
    expect(orphaned.isEmpty).toBe(true)
    expect(orphaned.labelAt('sh1', 0, 0)).toBeNull()
  })

  it('shares one empty index for a workbook that withholds nothing', () => {
    expect(NO_REDACTIONS.isEmpty).toBe(true)
    // A reader with no index ref at all must behave exactly as it did before.
    const bare = demoContext({}, undefined)
    delete bare.redactionIndexRef
    expect(redactionsOf(bare)).toBe(NO_REDACTIONS)
  })

  it('counts a rectangle of withheld cells without double-counting overlap', () => {
    const overlapping = buildRedactionIndex(
      [
        {
          sheetName: 'Customers',
          marks: [
            { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1, label: 'a' },
            { startRow: 2, endRow: 5, startColumn: 1, endColumn: 1, label: 'b' },
          ],
        },
      ],
      SHEETS,
    )
    const marks = overlapping.marksFor('sh1')
    // rows 0-3 x cols 0-1 = 8, plus rows 4-5 x col 1 = 2; the two-cell overlap
    // at rows 2-3 col 1 is inside the first rectangle already.
    expect(countWithheldIn(marks, { startRow: 0, endRow: 5, startColumn: 0, endColumn: 1 })).toBe(
      10,
    )
  })
})

describe('the label conventions shared with the other apps', () => {
  it('strips the characters that would break a carrier or a prompt', () => {
    expect(sanitizeLabel('  a{b}c<d>e="f"\'g*h/i  ')).toBe('abcdefghi')
  })

  it('renders the same {{label}} the other apps show', () => {
    expect(placeholderSource('客户电话')).toBe('{{客户电话}}')
    expect(placeholderSource('   ')).toBe('{{private}}')
  })
})

describe('read_cells projects every field of a withheld cell', () => {
  it('replaces the value and drops the formula and the raw value', () => {
    // Projecting only `value` would hand over the number twice over: `formula`
    // is the cell's own text and `rawValue` the model value behind it.
    const ctx = demoContext({
      B2: { value: '13800138000', formula: '=A1', rawValue: 13800138000 },
      B3: { value: 'public' },
    })
    const result = readCells(ctx, ['B2', 'B3'])
    expect(result.B2).toEqual({ value: '{{客户电话}}' })
    expect(result.B2?.formula).toBeUndefined()
    expect(result.B2?.rawValue).toBeUndefined()
    expect(result.B3).toEqual({ value: 'public' })
  })

  it('leaves a workbook with no marks exactly as it was', () => {
    const ctx = demoContext({ B2: { value: '13800138000' } }, NO_REDACTIONS)
    expect(readCells(ctx, ['B2']).B2).toEqual({ value: '13800138000' })
  })
})

describe('aggregate_range cannot be back-solved through a withheld cell', () => {
  /**
   * The threat: `sum` is exact, so a single withheld cell inside a column is
   * recoverable — total minus everything the model can see. Masking the cell
   * in the output does nothing, because the cell is not the only place its
   * value appears. The only fix is to keep it out of the arithmetic.
   *
   * The contract these tests pin is the aggregator's half of that: a withheld
   * cell is reported through addWithheld and its value never reaches add. The
   * routing half — that aggregateWorkbookRange actually takes that path for
   * every one of its three cell sources — is covered by aggregate-range.test.ts.
   */
  it('leaves the withheld value out of every statistic', () => {
    const aggregator = createRangeAggregator()
    for (const value of [10, 20, 40]) aggregator.add(value)
    aggregator.addWithheld(1) // the withheld cell: counted, value never passed in
    const result = aggregator.finish(10)

    expect(result.sum).toBe(70)
    expect(result.min).toBe(10)
    expect(result.max).toBe(40)
    expect(result.average).toBeCloseTo(70 / 3)
    expect(result.numericCount).toBe(3)
    expect(result.distinct).toBe(3)
  })

  it('keeps a withheld text value out of the frequency table', () => {
    // `topValues` is a direct disclosure, not an inference: a withheld name
    // appearing with a count hands the value over outright.
    const aggregator = createRangeAggregator()
    for (const value of ['a', 'b', 'a', 'b']) aggregator.add(value)
    aggregator.addWithheld(7) // seven cells all reading "Acme Ltd"
    const result = aggregator.finish(10)
    expect(result.topValues).toEqual([
      { value: 'a', count: 2 },
      { value: 'b', count: 2 },
    ])
    expect(JSON.stringify(result)).not.toContain('Acme')
  })

  it('still counts the cell, so the total is not quietly short', () => {
    const aggregator = createRangeAggregator()
    aggregator.add(10)
    aggregator.addWithheld(1)
    const result = aggregator.finish(10)
    // `cells` is the range's size; reporting 1 when 2 were read would make the
    // model present an incomplete total as if it were whole.
    expect(result.cells).toBe(2)
    expect(result.withheld).toBe(1)
    expect(result.nonEmpty).toBe(1)
  })

  it('says so in the tool output, before any number', () => {
    const aggregator = createRangeAggregator()
    aggregator.add(10)
    aggregator.addWithheld(1)
    const text = formatRangeAggregate('B1:B2', aggregator.finish(10), 10)
    expect(text).toContain('withheld: 1')
    expect(text.indexOf('withheld: 1')).toBeLessThan(text.indexOf('sum:'))
  })

  it('says nothing about withholding when there is none', () => {
    const aggregator = createRangeAggregator()
    aggregator.add(10)
    expect(formatRangeAggregate('B1:B1', aggregator.finish(10), 10)).not.toContain('withheld')
  })

  it('ignores a non-positive count rather than skewing the range size', () => {
    const aggregator = createRangeAggregator()
    aggregator.add(10)
    aggregator.addWithheld(0)
    aggregator.addWithheld(-3)
    expect(aggregator.finish(10)).toMatchObject({ cells: 1, withheld: 0 })
  })
})

describe('read_cells projects a withheld cell on the streaming path too', () => {
  it('drops formula and rawValue for a cell read through the lazy grid', async () => {
    // A real workbook is read through the lazy grid, not the in-memory
    // snapshot, so this path has its own branch in readCells. Projecting only
    // `value` here would hand over the number twice: once in `formula` and
    // once in `rawValue`.
    vi.resetModules()
    vi.doMock('../src/renderer/univer-sync', () => ({
      lazyCellReader: () => (address: string) => ({
        value: address === 'B2' ? '13800138000' : 'public',
        formula: address === 'B2' ? '=A1&""' : undefined,
        rawValue: address === 'B2' ? 13800138000 : 'public',
      }),
    }))
    const { readCells: readCellsLazy } = await import('../src/renderer/ai/workbook-readers')

    const worksheet = {
      getSheetId: () => 'sh1',
      getSheetName: () => 'Customers',
      getMergedRanges: () => [],
    }
    const ctx: WorkbookReadContext = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getActiveSheet: () => worksheet,
              getSheetBySheetId: () => worksheet,
              getActiveRange: () => null,
              getSheets: () => [worksheet],
            }),
          },
        },
      } as never,
      // any non-null state selects the streaming branch
      lazyWorkbookRef: { current: {} as never },
      adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } } as never,
      redactionIndexRef: { current: index },
    }

    const result = readCellsLazy(ctx, ['B2', 'B3'])
    expect(result.B2).toEqual({ value: '{{客户电话}}' })
    expect(result.B2?.formula).toBeUndefined()
    expect(result.B2?.rawValue).toBeUndefined()
    expect(result.B3).toEqual({ value: 'public', rawValue: 'public' })
    vi.doUnmock('../src/renderer/univer-sync')
    vi.resetModules()
  })
})

describe('the prompt that explains the placeholders', () => {
  it('names the labels actually in play', () => {
    const text = placeholderInstruction(['客户电话', '客户电话', '身份证号'])
    expect(text).toContain('- {{客户电话}}')
    expect(text).toContain('- {{身份证号}}')
    // Listed once each: a duplicate reads as two different placeholders.
    expect(text.match(/- \{\{客户电话\}\}/g)).toHaveLength(1)
  })

  it('warns about the two ways a model could recover a withheld value', () => {
    const text = placeholderInstruction(['x'])
    // Statistics: the model must not add the hidden numbers back to a total.
    expect(text).toContain('left out of every statistic')
    // Search: an empty result must not read as "not in the workbook".
    expect(text).toContain('absent from `find_cells` results')
  })
})
