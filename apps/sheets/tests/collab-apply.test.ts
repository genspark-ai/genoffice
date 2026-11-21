import { describe, expect, it } from 'vitest'

import { applyRemoteWorkbookSnapshot } from '../src/renderer/collab-apply'

/**
 * The snapshot apply hands visuals back to the caller (charts are not Univer
 * cells and cannot ride the command listener). These tests pin that contract
 * without needing a live Univer grid: a null runtime applies nothing but still
 * returns the embedded chart list.
 */
describe('applyRemoteWorkbookSnapshot', () => {
  it('returns the charts embedded in the snapshot', () => {
    const raw = JSON.stringify({
      revision: 1,
      sheets: [
        {
          id: 's1',
          name: 'S1',
          cells: {},
          visuals: [
            { id: 'chart-1', sheetId: 's1', kind: 'chart' },
            { id: 'chart-2', sheetId: 's1', kind: 'chart' },
          ],
        },
        { id: 's2', name: 'S2', cells: {}, visuals: [{ id: 'chart-3', sheetId: 's2', kind: 'chart' }] },
      ],
    })
    const { visuals } = applyRemoteWorkbookSnapshot(null, raw)
    expect(visuals.map((v) => v.id)).toEqual(['chart-1', 'chart-2', 'chart-3'])
  })

  it('returns no visuals for an empty or malformed snapshot', () => {
    expect(applyRemoteWorkbookSnapshot(null, 'not json').visuals).toEqual([])
    expect(applyRemoteWorkbookSnapshot(null, JSON.stringify({ revision: 1, sheets: [] })).visuals).toEqual([])
  })
})
