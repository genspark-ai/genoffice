import { describe, expect, it, vi } from 'vitest'
import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'
import { proposeOperations, type PlanContext } from '../src/renderer/plan-operations'

/// Demo-branch coverage of the AI-path quadratic-formula guard (the lazy
/// branch shares quadraticFormulaError, covered by formula-cost.test.ts; its
/// wiring needs a live Univer runtime, exercised by the real-app driver).

function demoContext(): PlanContext {
  const adapter = new InMemoryWorkbookAdapter({
    revision: 0,
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet1',
        // extent 5000 rows so whole-column criteria clamp past the limit
        cells: { A1: { value: 'x' }, A5000: { value: 'y' } },
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

describe('proposeOperations quadratic formula guard (demo workbook)', () => {
  it('rejects a distinct-count COUNTIF formula via set_formula', () => {
    const outcome = proposeOperations(
      demoContext(),
      [
        {
          op: 'set_formula',
          sheetId: 'sheet-1',
          address: 'B1',
          formula: '=SUMPRODUCT(1/COUNTIF(A1:A5000,A1:A5000))',
        },
      ],
      'distinct count',
    )
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('aggregate_range')
  })

  it('rejects quadratic formulas smuggled in through set_range strings', () => {
    const outcome = proposeOperations(
      demoContext(),
      [
        {
          op: 'set_range',
          sheetId: 'sheet-1',
          start: 'B1',
          values: [['label', '=COUNTIF(A:A,A:A)']],
        },
      ],
      'range write',
    )
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('element comparisons')
  })

  it('allows linear formulas over the same range', () => {
    const outcome = proposeOperations(
      demoContext(),
      [
        {
          op: 'set_formula',
          sheetId: 'sheet-1',
          address: 'B1',
          formula: '=SUM(A1:A5000)',
        },
      ],
      'sum',
    )
    expect(outcome.ok).toBe(true)
  })
})
