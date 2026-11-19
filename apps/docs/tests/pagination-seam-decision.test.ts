// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import {
  markTableSeamSlices,
  ROW_SEAM_TOL_PX,
  seamDecision,
  tableRowTops,
} from '../src/renderer/pagination-slices'
import { rowSplitCss } from '../src/renderer/pagination-lines'
import { leadLeakCss } from '../src/renderer/components/PaginationPreview'
import type { BlockBox, PageSlice } from '../src/renderer/pagination-types'

const table = (top: number, heights: number[]): BlockBox => {
  const el = document.createElement('table')
  el.dataset.idx = '1'
  const tbody = document.createElement('tbody')
  for (const h of heights) {
    const tr = document.createElement('tr')
    tr.innerHTML = `<td><p>${h}</p></td>`
    tbody.appendChild(tr)
  }
  el.appendChild(tbody)
  document.body.appendChild(el)
  const height = heights.reduce((a, b) => a + b, 0)
  return { top, height, el, tableRows: heights.map((h) => ({ height: h })) }
}

const para = (top: number, height: number): BlockBox => ({
  el: document.createElement('p'),
  top,
  height,
})

describe('seamDecision', () => {
  const tops = [100, 300, 500]

  it('returns the boundary the cut sits on and the signed drift', () => {
    expect(seamDecision(tops, 300)).toEqual({ row: 1, offset: 0 })
    expect(seamDecision(tops, 300.4)).toEqual({ row: 1, offset: expect.closeTo(0.4, 6) })
    expect(seamDecision(tops, 499.6)).toEqual({ row: 2, offset: expect.closeTo(-0.4, 6) })
  })

  it('treats a cut past the tolerance as mid-row', () => {
    expect(seamDecision(tops, 300 + ROW_SEAM_TOL_PX + 0.01)).toBeNull()
    expect(seamDecision(tops, 350)).toBeNull()
    expect(seamDecision([], 100)).toBeNull()
  })

  it('picks the nearest boundary across a hairline row', () => {
    expect(seamDecision([700, 720, 721], 720.6)).toEqual({
      row: 2,
      offset: expect.closeTo(-0.4, 6),
    })
  })

  it('tableRowTops lists every row top starting at the table top', () => {
    expect(tableRowTops({ top: 10, tableRows: [{ height: 20 }, { height: 30 }] })).toEqual([10, 30])
    expect(tableRowTops({ top: 10 })).toEqual([])
  })
})

describe('one seam decision across the engine flag, the preview lead row and row-split pages', () => {
  const cutAt = (y: number) => {
    const t = table(100, [200, 200, 100])
    const blocks = [para(0, 100), t, para(600, 100)]
    const slices: PageSlice[] = [
      { start: 0, end: y, section: 0 },
      { start: y, end: 700, section: 0 },
    ]
    markTableSeamSlices(slices, blocks)
    const leadCss = leadLeakCss(blocks, slices)
    return { cutTable: slices[1].cutTable === true, leadRowHidden: leadCss.includes('pv-lead') }
  }

  it('a cut within the tolerance of a row top is a seam for both', () => {
    expect(cutAt(300.4)).toEqual({ cutTable: true, leadRowHidden: true })
  })

  it('a cut drifting past the tolerance is mid-row for both', () => {
    expect(cutAt(300.6)).toEqual({ cutTable: false, leadRowHidden: false })
  })

  it('a split fragment starting on the seam lands on the page opening there', () => {
    const t = table(100, [300])
    const css = rowSplitCss(
      [
        {
          blockTop: 100,
          row: 0,
          rules: [
            { from: 200 - ROW_SEAM_TOL_PX, cell: 0, child: 0, tail: true, hide: true },
            { from: 200 - ROW_SEAM_TOL_PX - 0.01, cell: 0, child: 0, tail: true, hide: true },
          ],
        },
      ],
      [t],
      [{ start: 0 }, { start: 200 }],
    )
    expect(css.split('\n').map((r) => r.slice(0, 26))).toEqual([
      '.pv-page[data-pv-page="1"]',
      '.pv-page[data-pv-page="0"]',
    ])
  })
})
