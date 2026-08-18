import { describe, expect, it } from 'vitest'
import { offsetFormulaRefs } from '../src/domain/formula-shift'
import {
  expandToPrimitiveOps,
  MAX_EXPANDED_CELL_OPS,
  type WorkbookCommandBatch,
} from '../src/domain/workbook-dsl'
import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'
import { buildLazyChangePlan } from '../src/renderer/lazy-plan'
import { fillFormulaCostError } from '../src/renderer/formula-cost'

describe('offsetFormulaRefs (fill/copy reference semantics)', () => {
  it('shifts relative references by the copy offset', () => {
    expect(offsetFormulaRefs('=A2+1', 3, 0)).toBe('=A5+1')
    expect(offsetFormulaRefs('=SUM(A2:C2)', 1, 0)).toBe('=SUM(A3:C3)')
    expect(offsetFormulaRefs('=B2*C2', 0, 2)).toBe('=D2*E2')
  })

  it('pins $-anchored axes and shifts only the free ones', () => {
    expect(offsetFormulaRefs('=$A$2+B2', 1, 1)).toBe('=$A$2+C3')
    expect(offsetFormulaRefs('=B$2/SUM($B$4:$B$8)', 5, 1)).toBe('=C$2/SUM($B$4:$B$8)')
    expect(offsetFormulaRefs('=$B2', 2, 3)).toBe('=$B4')
  })

  it('never rewrites text inside string literals', () => {
    expect(offsetFormulaRefs('=IF(A1="A1","A1",B1)', 1, 0)).toBe('=IF(A2="A1","A1",B2)')
  })

  it('shifts sheet-qualified references, quoted names included', () => {
    expect(offsetFormulaRefs("=Sheet2!A1+'My Sheet'!B2", 1, 0)).toBe("=Sheet2!A2+'My Sheet'!B3")
  })

  it('turns references pushed off the grid into #REF!', () => {
    expect(offsetFormulaRefs('=A1', -1, 0)).toBe('=#REF!')
    expect(offsetFormulaRefs('=A1', 0, -1)).toBe('=#REF!')
  })

  it('shifts whole-column spans on fill-right and leaves them alone on fill-down', () => {
    expect(offsetFormulaRefs('=SUM(B:B)', 0, 1)).toBe('=SUM(C:C)')
    expect(offsetFormulaRefs('=SUM($B:$B)', 0, 1)).toBe('=SUM($B:$B)')
    expect(offsetFormulaRefs('=SUM(B:B)', 5, 0)).toBe('=SUM(B:B)')
  })

  it('does not mangle function names that look like references', () => {
    expect(offsetFormulaRefs('=LOG10(A1)', 1, 0)).toBe('=LOG10(A2)')
  })
})

describe('expandToPrimitiveOps range-level ops', () => {
  const readCell = (): { value: null } => ({ value: null })

  it('passes fill_range through as a single range-level primitive', () => {
    const expanded = expandToPrimitiveOps(
      [{ op: 'fill_range', sheetId: 's1', source: 'A2', target: 'A2:A88588' }],
      readCell,
    )
    expect(expanded).toHaveLength(1)
    expect(expanded[0]?.op).toBe('fill_range')
  })

  it('rejects a target that the source does not tile exactly', () => {
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'fill_range', sheetId: 's1', source: 'A2:B2', target: 'A2:C10' }],
        readCell,
      ),
    ).toThrow(/whole multiple/)
  })

  it('rejects an overlapping source that is not the target top-left corner', () => {
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'fill_range', sheetId: 's1', source: 'A5', target: 'A2:A10' }],
        readCell,
      ),
    ).toThrow(/overlap/)
  })

  it('caps fill_range target at 200,000 cells and source at 2000', () => {
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'fill_range', sheetId: 's1', source: 'A1', target: 'A1:C100000' }],
        readCell,
      ),
    ).toThrow(/200,000/)
    expect(() =>
      expandToPrimitiveOps(
        [{ op: 'fill_range', sheetId: 's1', source: 'A1:A2001', target: 'A1:B4002' }],
        readCell,
      ),
    ).toThrow(/source covers more than 2000/)
  })

  it('expands small clear_range per cell but passes large ones through', () => {
    const small = expandToPrimitiveOps(
      [{ op: 'clear_range', sheetId: 's1', range: 'A1:A3' }],
      readCell,
    )
    expect(small).toHaveLength(3)
    expect(small.every((op) => op.op === 'clear_cell')).toBe(true)

    const large = expandToPrimitiveOps(
      [{ op: 'clear_range', sheetId: 's1', range: `A1:A${MAX_EXPANDED_CELL_OPS + 1}` }],
      readCell,
    )
    expect(large).toHaveLength(1)
    expect(large[0]?.op).toBe('clear_range')

    expect(() =>
      expandToPrimitiveOps([{ op: 'clear_range', sheetId: 's1', range: 'A1:C100000' }], readCell),
    ).toThrow(/200,000/)
  })

  it('allows format_range over a whole large column now', () => {
    const expanded = expandToPrimitiveOps(
      [{ op: 'format_range', sheetId: 's1', range: 'A1:A88588', format: { bold: true } }],
      readCell,
    )
    expect(expanded).toHaveLength(1)
    expect(expanded[0]?.op).toBe('format_range')
  })
})

describe('InMemoryWorkbookAdapter fill_range (demo mode)', () => {
  function adapter(): InMemoryWorkbookAdapter {
    return new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: {
            A2: { value: null, formula: '=B2*2' },
            B2: { value: 10 },
            B3: { value: 20 },
            B4: { value: 30 },
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
      summary: 'fill test',
      operations,
    }
  }

  it('fills a formula down with shifted relative references', () => {
    const workbook = adapter()
    const plan = workbook.plan(
      batch([{ op: 'fill_range', sheetId: 'sheet-1', source: 'A2', target: 'A2:A4' }]),
    )
    expect(plan.cellChanges).toHaveLength(3)
    expect(plan.cellChanges.map((change) => change.after.formula)).toEqual([
      '=B2*2',
      '=B3*2',
      '=B4*2',
    ])
    workbook.apply(plan)
    const sheet = workbook.getSnapshot().sheets[0]
    expect(sheet?.cells.A3?.formula).toBe('=B3*2')
    expect(sheet?.cells.A4?.formula).toBe('=B4*2')
  })

  it('copies plain values when the source has no formula', () => {
    const workbook = adapter()
    const plan = workbook.plan(
      batch([{ op: 'fill_range', sheetId: 'sheet-1', source: 'B2', target: 'D2:D3' }]),
    )
    expect(plan.cellChanges.map((change) => change.after.value)).toEqual([10, 10])
  })

  it('caps demo fills at the per-cell expansion limit', () => {
    const workbook = adapter()
    expect(() =>
      workbook.plan(
        batch([{ op: 'fill_range', sheetId: 'sheet-1', source: 'A2', target: 'A2:A5000' }]),
      ),
    ).toThrow(/limited to 2000/)
  })

  it('clears only existing cells for a large range-level clear', () => {
    const workbook = adapter()
    const plan = workbook.plan(
      batch([{ op: 'clear_range', sheetId: 'sheet-1', range: 'A1:B3000' }]),
    )
    // 6000-cell range, but only the 4 populated cells produce changes.
    expect(plan.cellChanges).toHaveLength(4)
    expect(plan.cellChanges.every((change) => change.after.value === null)).toBe(true)
  })
})

describe('buildLazyChangePlan range-level entries', () => {
  const readCell = (): { value: null } => ({ value: null })

  it('plans fill_range and large clear_range as labeled bulk entries', () => {
    const batch: WorkbookCommandBatch = {
      dslVersion: 1,
      transactionId: 'tx-lazy',
      baseRevision: 0,
      summary: 'bulk ops',
      operations: [
        { op: 'fill_range', sheetId: 's1', source: 'A2', target: 'A2:A88588' },
        { op: 'clear_range', sheetId: 's1', range: 'B2:B88588' },
      ],
    }
    const plan = buildLazyChangePlan(batch, readCell, () => 'Sheet1')
    expect(plan.cellChanges).toHaveLength(0)
    expect(plan.structuralChanges.map((change) => change.label)).toEqual([
      'Fill A2 → A2:A88588',
      'Clear B2:B88588',
    ])
  })
})

describe('fillFormulaCostError', () => {
  const sheets = [{ name: 'Sheet1', rows: 88588, columns: 8 }]

  it('allows row-local relative formulas across a whole column', () => {
    expect(fillFormulaCostError('=B2+1', 88587, 88586, 0, 'Sheet1', sheets)).toBeNull()
    expect(fillFormulaCostError('=DATE(2026,8,18)', 88587, 88586, 0, 'Sheet1', sheets)).toBeNull()
  })

  it('rejects formulas that re-scan a large absolute range per copy', () => {
    expect(fillFormulaCostError('=SUM(B$2:B$88588)', 88587, 88586, 0, 'Sheet1', sheets)).toMatch(
      /element operations/,
    )
    expect(
      fillFormulaCostError('=VLOOKUP(A2,D$2:E$88588,2,0)', 88587, 88586, 0, 'Sheet1', sheets),
    ).toMatch(/element operations/)
  })

  it('rejects anchored-start ranges that expand as the fill grows', () => {
    // Running total =SUM(B$2:B2): the source copy scans one cell, but the
    // last copy scans the whole column — quadratic in total.
    expect(fillFormulaCostError('=SUM(B$2:B2)', 88587, 88586, 0, 'Sheet1', sheets)).toMatch(
      /element operations/,
    )
    // Small fills of the same shape stay fine (1000 rows ≈ 1e6/2 ops).
    expect(fillFormulaCostError('=SUM(B$2:B2)', 1000, 999, 0, 'Sheet1', sheets)).toBeNull()
  })
})
