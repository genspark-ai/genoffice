import { describe, expect, it, vi } from 'vitest'

import { updateExternalLinksFlow } from '../src/renderer/external-links-flow'
import type { ExternalLinksUpdateResult } from '../src/shared/desktop-api'

const REPORT: ExternalLinksUpdateResult = {
  links: [
    {
      index: 1,
      rawTarget: 'source.xlsx',
      source: '/tmp/data/source.xlsx',
      reachable: true,
      sheetNames: ['Sheet1'],
    },
  ],
  cells: [
    {
      sheet: 'Sheet1',
      row: 0,
      column: 0,
      formatted: '20',
      number: 20,
      kind: 'number',
      formula: '[1]Sheet1!A1',
    },
  ],
}

describe('updateExternalLinksFlow', () => {
  it('saves first, applies the report, and saves again', async () => {
    const order: string[] = []
    const save = vi.fn(async () => {
      order.push('save')
      return { ok: true }
    })
    const update = vi.fn(async (path: string) => {
      order.push('update:' + path)
      return REPORT
    })
    const apply = vi.fn((report: ExternalLinksUpdateResult) => {
      order.push('apply:' + report.cells.length)
    })
    const statuses: ExternalLinksUpdateResult[] = []
    const updated = await updateExternalLinksFlow({
      path: '/tmp/book.xlsx',
      save,
      update,
      apply,
      onStatuses: (report) => {
        order.push('statuses')
        statuses.push(report)
      },
    })
    expect(updated).toBe(true)
    // the disk write precedes the read the report is computed from, and the
    // second save persists the applied values — never update-then-save-then-apply
    expect(order).toEqual(['save', 'update:/tmp/book.xlsx', 'apply:1', 'statuses', 'save'])
    expect(statuses).toHaveLength(1)
  })

  it('aborts before the update when the first save fails', async () => {
    const update = vi.fn(async () => REPORT)
    const apply = vi.fn()
    const updated = await updateExternalLinksFlow({
      path: '/tmp/book.xlsx',
      save: vi.fn(async () => ({ ok: false })),
      update,
      apply,
      onStatuses: vi.fn(),
    })
    expect(updated).toBe(false)
    expect(update).not.toHaveBeenCalled()
    expect(apply).not.toHaveBeenCalled()
  })

  it('applies and saves even when the report is empty', async () => {
    const save = vi.fn(async () => ({ ok: true }))
    const apply = vi.fn()
    const updated = await updateExternalLinksFlow({
      path: '/tmp/book.xlsx',
      save,
      update: vi.fn(async () => ({ links: [], cells: [] })),
      apply,
      onStatuses: vi.fn(),
    })
    expect(updated).toBe(true)
    expect(apply).toHaveBeenCalledWith({ links: [], cells: [] })
    expect(save).toHaveBeenCalledTimes(2)
  })
})
