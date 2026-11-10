import { describe, expect, it } from 'vitest'
import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'
import { computeSortedRowOrder } from '../src/domain/sort-range'

describe('computeSortedRowOrder non-finite keys', () => {
  it('sorts finite numbers deterministically around non-finite keys', () => {
    const rows = [[3], [NaN], [1], [Infinity], [2], [-Infinity]] as const
    const first = computeSortedRowOrder(
      rows.map((r) => [...r]),
      0,
      true,
    )
    // finite ascending first, then non-finite, blanks last per Excel order
    expect(first.slice(0, 3).map((i) => rows[i]![0])).toEqual([1, 2, 3])
    // repeated runs are identical (no NaN-comparator flakiness)
    for (let k = 0; k < 10; k++) {
      expect(
        computeSortedRowOrder(
          rows.map((r) => [...r]),
          0,
          true,
        ),
      ).toEqual(first)
    }
  })

  it('keeps normal numeric order unchanged', () => {
    const rows = [[3], [1], [2]] as const
    expect(
      computeSortedRowOrder(
        rows.map((r) => [...r]),
        0,
        true,
      ),
    ).toEqual([1, 2, 0])
    expect(
      computeSortedRowOrder(
        rows.map((r) => [...r]),
        0,
        false,
      ),
    ).toEqual([0, 2, 1])
  })
})

describe('sort_range over formatted cells', () => {
  it('guards each cell on its display value, so a currency sort plans', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: {
            A1: { value: 'Amount' },
            A2: { value: '$30.00', rawValue: 30 },
            A3: { value: '$10.00', rawValue: 10 },
            A4: { value: '$20.00', rawValue: 20 },
          },
        },
      ],
    })
    // The precondition used to carry the raw serial, which the CAS then
    // compared against the "$30.00" display text — the whole sort threw.
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-1',
      baseRevision: 0,
      summary: 'sort by amount',
      operations: [
        {
          op: 'sort_range',
          sheetId: 'sheet-1',
          range: 'A1:A4',
          byColumn: 'A',
          order: 'asc',
          hasHeader: true,
        },
      ],
    })
    // Raw values move, so the number format is not corrupted into text.
    expect(plan.cellChanges.map((change) => change.after.value)).toEqual([10, 20, 30])
  })
})
