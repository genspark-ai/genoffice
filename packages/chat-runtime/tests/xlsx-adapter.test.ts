/**
 * Cross-check that the XLSX adapter is shape-compatible with the
 * chat-runtime's `normalizeChangePlan`. The adapter itself lives in
 * apps/sheets/src/renderer/ai/xlsx-change-plan.ts and is exercised by
 * the XLSX test suite; this file verifies the contract from the runtime
 * side so a runtime regression can't silently break the XLSX preview.
 */
import { describe, expect, it } from 'vitest'

import { normalizeChangePlan, summarizeChangePlan } from '../src/change-plan.js'
import type { ChatChangePlanOp } from '../src/types.js'

describe('XLSX ChangePlan shape', () => {
  it('accepts a flat array of XLSX WorkbookOperations', () => {
    const ops = [
      { op: 'insert-column', at: 'C' },
      { op: 'set-cell', at: 'C1', value: 'Total' },
      { op: 'set-formula', at: 'C2', formula: '=SUM(B2:B10)' },
    ]
    const plan = normalizeChangePlan({ app: 'sheets', title: 'Insert column', ops })
    expect(plan.ops).toHaveLength(3)
    for (const op of plan.ops) {
      expect((op as { kind: string }).kind).toBe('workbook')
    }
  })

  it('summarizeChangePlan renders one bullet per XLSX op', () => {
    const plan = normalizeChangePlan({
      app: 'sheets',
      title: 'Insert column',
      ops: [{ op: 'insert-column', at: 'C' }, { op: 'set-cell', at: 'C1', value: 'Total' }],
    })
    const bullets = summarizeChangePlan(plan, op => {
      const inner = (op as ChatChangePlanOp).ops?.[0] as { at?: string; value?: unknown } | undefined
      if (!inner) return '[workbook]'
      if (inner.value !== undefined) return `${inner.at} = ${String(inner.value)}`
      return `${(inner as { at: string }).at}`
    })
    expect(bullets).toHaveLength(2)
    expect(bullets[0]).toBe('C')
    expect(bullets[1]).toBe('C1 = Total')
  })

  it('preserves warnings through the adapter contract', () => {
    const plan = normalizeChangePlan({
      app: 'sheets',
      title: 'Risky change',
      ops: [{ op: 'delete-sheet', at: 'B' }],
      warnings: ['deleting a sheet is irreversible'],
      requireConfirm: true,
    })
    expect(plan.warnings).toEqual(['deleting a sheet is irreversible'])
    expect(plan.requireConfirm).toBe(true)
  })
})
