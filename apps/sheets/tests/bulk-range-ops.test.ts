import { describe, expect, it } from 'vitest'
import {
  copyTargetBounds,
  expandToPrimitiveOps,
  MAX_EXPANDED_CELL_OPS,
  replaceOccurrences,
  type WorkbookCommandBatch,
} from '../src/domain/workbook-dsl'
import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'
import { buildLazyChangePlan } from '../src/renderer/lazy-plan'

const readCell = (): { value: null } => ({ value: null })

describe('copy_range validation and expansion', () => {
  it('passes through as a single range-level primitive', () => {
    const expanded = expandToPrimitiveOps(
      [{ op: 'copy_range', sheetId: 's1', source: 'A1:F5000', target: 'H1' }],
      readCell,
    )
    expect(expanded).toHaveLength(1)
    expect(expanded[0]?.op).toBe('copy_range')
  })

  it('extends a single-cell target to the source size', () => {
    expect(
      copyTargetBounds({ op: 'copy_range', sheetId: 's1', source: 'B2:D4', target: 'H10' }),
    ).toEqual({ startRow: 9, endRow: 11, startColumn: 7, endColumn: 9 })
  })

  it('accepts a full-size target range but rejects a mismatched one', () => {
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'copy_range', sheetId: 's1', source: 'B2:D4', target: 'H10:J12' }],
        readCell,
      ),
    ).not.toThrow()
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'copy_range', sheetId: 's1', source: 'B2:D4', target: 'H10:K12' }],
        readCell,
      ),
    ).toThrow(/does not match the source's size/)
  })

  it('rejects overlapping source and target on the same sheet', () => {
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'copy_range', sheetId: 's1', source: 'A1:C10', target: 'B5' }],
        readCell,
      ),
    ).toThrow(/overlap/)
    // Same coordinates on another sheet are fine.
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'copy_range', sheetId: 's2', sourceSheetId: 's1', source: 'A1:C10', target: 'A1' }],
        readCell,
      ),
    ).not.toThrow()
  })

  it('caps the source at 200,000 cells', () => {
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'copy_range', sheetId: 's1', source: 'A1:C100000', target: 'H1' }],
        readCell,
      ),
    ).toThrow(/200,000/)
  })
})

describe('convert_to_values expansion', () => {
  it('passes through as a single range-level primitive', () => {
    const expanded = expandToPrimitiveOps(
      [{ op: 'convert_to_values', sheetId: 's1', range: 'B2:B88588' }],
      readCell,
    )
    expect(expanded).toHaveLength(1)
    expect(expanded[0]?.op).toBe('convert_to_values')
  })

  it('caps the range at 200,000 cells', () => {
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'convert_to_values', sheetId: 's1', range: 'A1:C100000' }],
        readCell,
      ),
    ).toThrow(/200,000/)
  })
})

describe('convert_to_values batch-order guard', () => {
  it('rejects same-batch formula writes into the convert range', () => {
    expect(() =>
      expandToPrimitiveOps(
        [
          { op: 'set_formula', sheetId: 's1', address: 'B5', formula: '=A5*2' },
          { op: 'convert_to_values', sheetId: 's1', range: 'B2:B10' },
        ],
        readCell,
      ),
    ).toThrow(/cannot share a batch/)
    expect(() =>
      expandToPrimitiveOps(
        [
          { op: 'set_range', sheetId: 's1', start: 'B2', values: [['=A2*2']] },
          { op: 'convert_to_values', sheetId: 's1', range: 'B2:B10' },
        ],
        readCell,
      ),
    ).toThrow(/cannot share a batch/)
    expect(() =>
      expandToPrimitiveOps(
        [
          { op: 'fill_range', sheetId: 's1', source: 'B2', target: 'B2:B10' },
          { op: 'convert_to_values', sheetId: 's1', range: 'B2:B10' },
        ],
        readCell,
      ),
    ).toThrow(/cannot share a batch/)
  })

  it('allows non-overlapping writes and plain-value overlaps', () => {
    expect(() =>
      expandToPrimitiveOps(
        [
          { op: 'set_formula', sheetId: 's1', address: 'D1', formula: '=A1' },
          { op: 'convert_to_values', sheetId: 's1', range: 'B2:B10' },
        ],
        readCell,
      ),
    ).not.toThrow()
    // Overlap on another sheet is fine.
    expect(() =>
      expandToPrimitiveOps(
        [
          { op: 'set_formula', sheetId: 's2', address: 'B5', formula: '=A5' },
          { op: 'convert_to_values', sheetId: 's1', range: 'B2:B10' },
        ],
        readCell,
      ),
    ).not.toThrow()
    // Plain values commute with the convert (it only touches formula cells).
    expect(() =>
      expandToPrimitiveOps(
        [
          { op: 'set_range', sheetId: 's1', start: 'B2', values: [[42]] },
          { op: 'convert_to_values', sheetId: 's1', range: 'B2:B10' },
        ],
        readCell,
      ),
    ).not.toThrow()
  })
})

describe('find_replace large-range expansion', () => {
  it('still expands small ranges into per-cell edits', () => {
    const expanded = expandToPrimitiveOps(
      [{ op: 'find_replace', sheetId: 's1', range: 'A1:A3', find: 'x', replace: 'y' }],
      (address) => (address === 'A2' ? { value: 'x marks' } : { value: null }),
    )
    expect(expanded).toEqual([{ op: 'set_cell', sheetId: 's1', address: 'A2', value: 'y marks' }])
  })

  it('passes ranges above the expansion cap through range-level, without reading cells', () => {
    const expanded = expandToPrimitiveOps([
      {
        op: 'find_replace',
        sheetId: 's1',
        range: `A1:A${MAX_EXPANDED_CELL_OPS + 1}`,
        find: 'x',
        replace: 'y',
      },
    ])
    expect(expanded).toHaveLength(1)
    expect(expanded[0]?.op).toBe('find_replace')
  })

  it('caps the range at 200,000 cells', () => {
    expect(() =>
      expandToPrimitiveOps([
        { op: 'find_replace', sheetId: 's1', range: 'A1:C100000', find: 'x', replace: 'y' },
      ]),
    ).toThrow(/200,000/)
  })
})

describe('replaceOccurrences (shared by expansion and the chunked executor)', () => {
  it('replaces all occurrences, case-insensitive by default', () => {
    expect(replaceOccurrences('Foo foo FOO', 'foo', 'bar', false)).toBe('bar bar bar')
    expect(replaceOccurrences('Foo foo FOO', 'foo', 'bar', true)).toBe('Foo bar FOO')
  })

  it('keeps a literal $ in the replacement literal', () => {
    expect(replaceOccurrences('price: X', 'X', '$1.00', false)).toBe('price: $1.00')
  })
})

describe('InMemoryWorkbookAdapter (demo mode)', () => {
  function adapter(): InMemoryWorkbookAdapter {
    return new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: {
            A1: { value: 'Header' },
            A2: { value: null, formula: '=B2*2' },
            B2: { value: 10 },
            B3: { value: 20 },
          },
        },
      ],
    })
  }

  function batch(operations: unknown[]): unknown {
    return {
      dslVersion: 1,
      transactionId: `tx-${Math.random().toString(36).slice(2)}`,
      baseRevision: 0,
      summary: 'bulk ops test',
      operations,
    }
  }

  it('copies a block once with shifted relative references', () => {
    const workbook = adapter()
    const plan = workbook.plan(
      batch([{ op: 'copy_range', sheetId: 'sheet-1', source: 'A1:B2', target: 'D1' }]),
    )
    workbook.apply(plan)
    const sheet = workbook.getSnapshot().sheets[0]
    expect(sheet?.cells.D1?.value).toBe('Header')
    expect(sheet?.cells.D2?.formula).toBe('=E2*2')
    expect(sheet?.cells.E2?.value).toBe(10)
    // The source stays in place — copy, not move.
    expect(sheet?.cells.A1?.value).toBe('Header')
  })

  it('caps demo copies at the per-cell expansion limit', () => {
    const workbook = adapter()
    expect(() =>
      workbook.plan(
        batch([{ op: 'copy_range', sheetId: 'sheet-1', source: 'A1:A2001', target: 'D1' }]),
      ),
    ).toThrow(/limited to 2000/)
  })

  it('rejects range-level convert_to_values and find_replace with clear messages', () => {
    const workbook = adapter()
    expect(() =>
      workbook.plan(batch([{ op: 'convert_to_values', sheetId: 'sheet-1', range: 'A1:B3000' }])),
    ).toThrow(/convert_to_values/)
    expect(() =>
      workbook.plan(
        batch([
          { op: 'find_replace', sheetId: 'sheet-1', range: 'A1:B3000', find: 'x', replace: 'y' },
        ]),
      ),
    ).toThrow(/limited to 2000/)
  })
})

describe('buildLazyChangePlan range-level entries', () => {
  it('plans copy_range, convert_to_values, and large find_replace as labeled bulk entries', () => {
    const batch: WorkbookCommandBatch = {
      dslVersion: 1,
      transactionId: 'tx-lazy',
      baseRevision: 0,
      summary: 'bulk ops',
      operations: [
        { op: 'copy_range', sheetId: 's1', source: 'A1:F5000', target: 'H1' },
        { op: 'convert_to_values', sheetId: 's1', range: 'B2:B88588' },
        { op: 'find_replace', sheetId: 's1', range: 'A1:A88588', find: 'old', replace: 'new' },
      ],
    }
    const plan = buildLazyChangePlan(batch, readCell, () => 'Sheet1')
    expect(plan.cellChanges).toHaveLength(0)
    expect(plan.structuralChanges.map((change) => change.label)).toEqual([
      'Copy A1:F5000 → H1',
      'Convert B2:B88588 to values',
      'Replace "old" → "new" in A1:A88588',
    ])
  })
})
