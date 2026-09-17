import { describe, expect, it } from 'vitest'
import { EDGE_LEAK_PX, leadLeakCss } from '../src/renderer/components/PaginationPreview'
import type { BlockBox, PageSlice } from '../src/renderer/pagination'

const block = (
  idx: number,
  top: number,
  height: number,
  extra: Partial<BlockBox> = {},
): BlockBox => {
  const el = document.createElement('p')
  el.dataset.idx = String(idx)
  return { top, height, el, ...extra }
}

const slice = (start: number, end: number, extra: Partial<PageSlice> = {}): PageSlice => ({
  start,
  end,
  section: 0,
  ...extra,
})

const hidden = (css: string): Array<[number, number]> =>
  [
    ...css.matchAll(
      /data-pv-page="(\d+)"\] \.pv-content > \[data-idx="(\d+)"\]\{visibility:hidden;\}/g,
    ),
  ].map((m) => [Number(m[1]), Number(m[2])])

describe('leadLeakCss', () => {
  // a break-only paragraph (two line boxes) ends page 1; the heading it pushes
  // opens page 2 exactly at the window end, where its glyph tops would leak
  const pages = [slice(0, 739), slice(739, 1500), slice(1500, 2000)]
  const breakPara = block(3, 665, 43)
  const heading = block(4, 739, 21)
  const body = block(5, 760, 64)

  it('hides the next page lead block on the page before it only', () => {
    expect(hidden(leadLeakCss([breakPara, heading, body], pages))).toEqual([[0, 4]])
  })

  it('leaves blocks that straddle the window end visible', () => {
    const straddling = block(4, 700, 100)
    expect(hidden(leadLeakCss([breakPara, straddling], pages))).toEqual([])
  })

  it('ignores blocks too far below the window to reach it', () => {
    const far = block(4, 739 + EDGE_LEAK_PX, 21)
    expect(hidden(leadLeakCss([far], pages))).toEqual([])
  })

  it('never hides on the last page (its window opens to full capacity)', () => {
    const tail = block(9, 2000, 21)
    expect(hidden(leadLeakCss([tail], pages))).toEqual([])
  })

  it('leaves floated blocks and pinned-box carriers alone', () => {
    const floated = block(4, 739, 21, { floated: true })
    const carrier = block(4, 739, 21)
    carrier.el!.innerHTML = '<div data-pin-page="0"></div>'
    expect(hidden(leadLeakCss([floated, carrier], pages))).toEqual([])
  })

  it('skips column-region pages, whose windows are per column', () => {
    const regioned = [slice(0, 739, { regions: [] }), slice(739, 1500)]
    expect(hidden(leadLeakCss([heading], regioned))).toEqual([])
  })
})

describe('leadLeakCss table seams', () => {
  const table = (top: number, heights: number[]) => {
    const el = document.createElement('table')
    el.dataset.idx = '7'
    const tbody = document.createElement('tbody')
    for (const h of heights) {
      const tr = document.createElement('tr')
      tr.innerHTML = `<td><p>${h}</p></td>`
      tbody.appendChild(tr)
    }
    el.appendChild(tbody)
    document.body.appendChild(el)
    const height = heights.reduce((a, b) => a + b, 0)
    return { top, height, el, tableRows: heights.map((h) => ({ height: h })) } as BlockBox
  }

  it('hides the cell contents of the row opening the next page, keeping the row itself', () => {
    const t = table(700, [20, 30, 40])
    const css = leadLeakCss([t], [slice(0, 750), slice(750, 1500)])
    expect(css).toBe(
      '.pv-page[data-pv-page="0"] .pv-content tr[data-pv-lead="0"] > * > *{visibility:hidden;}',
    )
    expect(t.el!.querySelectorAll('tr')[2].dataset.pvLead).toBe('0')
  })

  it('leaves a row cut mid-way alone and clears stale stamps', () => {
    const t = table(700, [20, 30, 40])
    t.el!.querySelectorAll('tr')[1].dataset.pvLead = '9'
    expect(leadLeakCss([t], [slice(0, 735), slice(735, 1500)])).toBe('')
    expect(t.el!.querySelectorAll('tr')[1].dataset.pvLead).toBeUndefined()
  })
})

describe('leadLeakCss short rows', () => {
  it('stamps the row whose top is the window end, not a hairline row just above it', () => {
    const el = document.createElement('table')
    el.dataset.idx = '8'
    const tbody = document.createElement('tbody')
    for (const h of [20, 1, 40]) {
      const tr = document.createElement('tr')
      tr.innerHTML = `<td><p>${h}</p></td>`
      tbody.appendChild(tr)
    }
    el.appendChild(tbody)
    document.body.appendChild(el)
    const t = {
      top: 700,
      height: 61,
      el,
      tableRows: [{ height: 20 }, { height: 1 }, { height: 40 }],
    } as BlockBox
    leadLeakCss([t], [slice(0, 721), slice(721, 1500)])
    const trs = el.querySelectorAll('tr')
    expect(trs[1].dataset.pvLead).toBeUndefined()
    expect(trs[2].dataset.pvLead).toBe('0')
  })
})
