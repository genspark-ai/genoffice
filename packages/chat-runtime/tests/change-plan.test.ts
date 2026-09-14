/**
 * ChangePlan tests — covers normalising different op shapes and rendering
 * the default text summary.
 */
import { describe, expect, it } from 'vitest'

import { normalizeChangePlan, summarizeChangePlan } from '../src/change-plan.js'

describe('normalizeChangePlan', () => {
  it('handles empty ops', () => {
    const plan = normalizeChangePlan({ app: 'docs', title: 'empty', ops: [] })
    expect(plan.ops).toEqual([])
    expect(plan.summary).toBe('No operations')
  })

  it('wraps XLSX-style ops', () => {
    const plan = normalizeChangePlan({
      app: 'sheets',
      title: 'Apply formula',
      ops: [{ op: 'set-formula', cell: 'A1', formula: '=SUM(B1:B10)' }],
    })
    expect(plan.ops[0]).toMatchObject({ kind: 'workbook' })
  })
})

describe('summarizeChangePlan', () => {
  it('produces one bullet per op without a custom renderer', () => {
    const plan = normalizeChangePlan({
      app: 'docs',
      title: 'Insert heading',
      ops: [
        { kind: 'doc', ops: [{ op: 'insert-heading' }], description: 'Insert H1 at top' },
      ],
    })
    const bullets = summarizeChangePlan(plan)
    expect(bullets).toHaveLength(1)
    expect(bullets[0]).toContain('[doc]')
  })

  it('uses a custom renderer when provided', () => {
    const plan = normalizeChangePlan({ app: 'docs', title: 't', ops: [{ kind: 'doc', ops: [{}] }] })
    const bullets = summarizeChangePlan(plan, op => `custom:${op.kind}`)
    expect(bullets[0]).toBe('custom:doc')
  })
})
