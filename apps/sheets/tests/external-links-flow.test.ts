import { describe, expect, it, vi } from 'vitest'

import { breakExternalLinkFlow, updateExternalLinksFlow } from '../src/renderer/external-links-flow'
import type { ExternalLinksUpdateResult, ExternalLinkUsageResult } from '../src/shared/desktop-api'

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

const USAGE: ExternalLinkUsageResult = {
  links: [
    {
      index: 1,
      cells: [
        {
          sheet: 'Sheet1',
          row: 0,
          column: 0,
          value: { kind: 'number', value: 20 },
        },
        {
          sheet: 'Sheet1',
          row: 1,
          column: 0,
          value: { kind: 'text', value: 'gone!' },
        },
      ],
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
    const update = vi.fn(async () => {
      order.push('update')
      return REPORT
    })
    const apply = vi.fn((report: ExternalLinksUpdateResult) => {
      order.push('apply:' + report.cells.length)
    })
    const statuses: ExternalLinksUpdateResult[] = []
    const updated = await updateExternalLinksFlow({
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
    expect(order).toEqual(['save', 'update', 'apply:1', 'statuses', 'save'])
    expect(statuses).toHaveLength(1)
  })

  it('treats a no-edits first save as a no-op and still updates', async () => {
    // the panel's common case: an untouched linked workbook. handleSave
    // answers { ok: false, reason: 'no-edits' } because the recalc overlay is
    // not journal work — the disk already matches the grid, so proceed.
    const save = vi.fn(async () => ({ ok: false, reason: 'no-edits' }))
    const update = vi.fn(async () => REPORT)
    const apply = vi.fn()
    const updated = await updateExternalLinksFlow({
      save,
      update,
      apply,
      onStatuses: vi.fn(),
    })
    expect(updated).toBe(true)
    expect(update).toHaveBeenCalledTimes(1)
    expect(apply).toHaveBeenCalledWith(REPORT)
    expect(save).toHaveBeenCalledTimes(2)
  })

  it('aborts before the update when the first save fails', async () => {
    const update = vi.fn(async () => REPORT)
    const apply = vi.fn()
    const updated = await updateExternalLinksFlow({
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

describe('breakExternalLinkFlow', () => {
  it('journals the link cells, queues one break, and saves', async () => {
    const order: string[] = []
    const restore = vi.fn(async () => {
      order.push('restore')
    })
    const outcome = await breakExternalLinkFlow({
      index: 1,
      readUsage: async () => USAGE,
      journal: async (cells) => {
        order.push('journal:' + cells.length)
        return restore
      },
      queue: (ops) => order.push('queue:' + ops.broken.join(',')),
      clear: () => order.push('clear'),
      save: async () => {
        order.push('save')
        return { ok: true }
      },
    })
    expect(outcome).toBe(true)
    // the grid is journaled before the save that carries the break; a landed
    // save keeps the literals
    expect(order).toEqual(['journal:2', 'queue:1', 'save', 'clear'])
    expect(restore).not.toHaveBeenCalled()
  })

  it('puts the formulas back when the save does not land', async () => {
    const order: string[] = []
    const outcome = await breakExternalLinkFlow({
      index: 1,
      readUsage: async () => USAGE,
      journal: async () => async () => {
        order.push('restore')
      },
      queue: () => {},
      clear: () => order.push('clear'),
      save: async () => ({ ok: false }),
    })
    expect(outcome).toBe(false)
    // the file still has the formulas and the link, so the grid must too
    expect(order).toEqual(['restore', 'clear'])
  })

  it('still queues the break when the usage scan lists no cells', async () => {
    // a link referenced only from defined names journals nothing; the file
    // transform still has to run (and may refuse)
    const journal = vi.fn(async () => async () => {})
    const queue = vi.fn()
    const save = vi.fn(async () => ({ ok: true }))
    const outcome = await breakExternalLinkFlow({
      index: 1,
      readUsage: async () => ({ links: [] }),
      journal,
      queue,
      clear: vi.fn(),
      save,
    })
    expect(outcome).toBe(true)
    expect(journal).not.toHaveBeenCalled()
    expect(queue).toHaveBeenCalledWith({ broken: [1], retargets: [] })
    expect(save).toHaveBeenCalledTimes(1)
  })
})
