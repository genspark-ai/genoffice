import { describe, expect, it } from 'vitest'
import {
  GAP_BAND,
  alignGapHfStrips,
  makeGapEl,
  syncCutOverlays,
  syncPhantomRowspans,
  clampCellBoxTops,
} from '../src/renderer/editor/pagination-gaps'
import { createLineRectsCache, singleCutCell } from '../src/renderer/pagination'

const rowOf = (cells: number): HTMLTableRowElement => {
  const tr = document.createElement('tr')
  for (let i = 0; i < cells; i++) tr.appendChild(document.createElement('td'))
  return tr
}

const m = { marginTop: 96, marginBottom: 96, marginLeft: 90, marginRight: 90 }

describe('in-row table cut decorations', () => {
  it('single-cell row: real inline gap band, not a zero-height cut marker', () => {
    const tr = rowOf(1)
    expect(singleCutCell(tr)).toBe(tr.firstElementChild)
    const el = makeGapEl(m, 'cell')
    // page-gap-inline is what the measurement gap-subtraction keys on
    expect(el.classList.contains('page-gap')).toBe(true)
    expect(el.classList.contains('page-gap-inline')).toBe(true)
    expect(el.classList.contains('page-gap-cut')).toBe(false)
    expect(el.style.height).toBe(`${96 + GAP_BAND + 96}px`)
    expect(el.style.getPropertyValue('--gap-mb')).toBe('96px')
    expect(el.style.width).toBe('calc(100% + 180px)')
  })

  it('multi-cell rows without an anchor / missing rows keep the zero-height cut marker', () => {
    expect(singleCutCell(rowOf(2))).toBeNull()
    expect(singleCutCell(null)).toBeNull()
    expect(makeGapEl(m, 'cut').className).toBe('page-gap-cut')
  })

  it('table gap row spans exactly the host grid: phantom columns collapse colgroup-less fixed-layout tables', () => {
    const gap = makeGapEl(m, 'table', 3)
    expect(gap.tagName).toBe('TR')
    const cell = gap.firstElementChild as HTMLTableCellElement
    expect(cell.tagName).toBe('TD')
    expect(cell.colSpan).toBe(3)
    // no count supplied: never widen the grid
    const bare = makeGapEl(m, 'table').firstElementChild as HTMLTableCellElement
    expect(bare.colSpan).toBe(1)
  })

  const rectOf = (top: number, height: number) =>
    ({ top, bottom: top + height, height, width: 100 }) as DOMRect
  const cellAnchorRow = (siblingBottom: number) => {
    // host cell: article content with the cut anchor at y=500; sibling: spacer content
    const tr = rowOf(2)
    const host = tr.children[0] as HTMLElement
    const p = document.createElement('p')
    const text = document.createTextNode('article body')
    p.appendChild(text)
    p.getBoundingClientRect = () => rectOf(500, 20)
    host.appendChild(p)
    const sibling = tr.children[1] as HTMLElement
    const sp = document.createElement('p')
    sp.getBoundingClientRect = () => rectOf(0, siblingBottom)
    sibling.appendChild(sp)
    return { tr, anchor: { node: text, charOffset: 0 } }
  }

  it('multi-cell row hosts a real band when sibling content ends above the cut', () => {
    const { tr, anchor } = cellAnchorRow(4) // 1px-spacer-gif sliver cell
    expect(singleCutCell(tr, anchor)).toBe(tr.children[0])
  })

  it('multi-cell row keeps the cut marker when a sibling has content below the cut', () => {
    const { tr, anchor } = cellAnchorRow(900)
    expect(singleCutCell(tr, anchor)).toBeNull()
  })

  it('a sibling block holding only a spacer strut below the cut still hosts a band', () => {
    const { tr, anchor } = cellAnchorRow(900)
    const sp = tr.children[1].firstElementChild as HTMLElement
    const strut = document.createElement('img')
    // 1px-wide vertical strut gif reaching below the cut (the sample-72 spacer)
    strut.getBoundingClientRect = () => ({ top: 0, bottom: 900, height: 900, width: 1 }) as DOMRect
    sp.appendChild(strut)
    expect(singleCutCell(tr, anchor)).toBe(tr.children[0])
  })
})

describe('phantom-row rowspan bridging', () => {
  const rootOf = (rows: string): HTMLElement => {
    const root = document.createElement('div')
    root.innerHTML = `<table><tbody>${rows}</tbody></table>`
    return root
  }
  const cell = (root: HTMLElement, id: string) =>
    root.querySelector(`#${id}`) as HTMLTableCellElement

  it('grows a rowspan crossing the insertion point and records the base', () => {
    const root = rootOf(`
      <tr><td id="a" rowspan="3"></td><td></td></tr>
      <tr><td></td></tr>
      <tr class="page-gap"><td colspan="1000"></td></tr>
      <tr><td></td></tr>
      <tr><td id="b" rowspan="2"></td><td></td></tr>
      <tr><td></td></tr>`)
    syncPhantomRowspans(root)
    expect(cell(root, 'a').rowSpan).toBe(4)
    expect(cell(root, 'a').getAttribute('data-base-rowspan')).toBe('3')
    // span entirely below the gap is untouched
    expect(cell(root, 'b').rowSpan).toBe(2)
    expect(cell(root, 'b').hasAttribute('data-base-rowspan')).toBe(false)
  })

  it('accumulates per insertion point, ignores spans ending at the boundary, restores on removal', () => {
    const root = rootOf(`
      <tr><td id="a" rowspan="5"></td><td id="c" rowspan="2"></td><td></td></tr>
      <tr><td></td></tr>
      <tr class="page-gap"><td></td></tr>
      <tr class="page-repeat-header"><td></td><td></td><td></td></tr>
      <tr><td></td><td></td></tr>
      <tr><td></td><td></td></tr>
      <tr class="page-gap"><td></td></tr>
      <tr><td></td><td></td></tr>`)
    syncPhantomRowspans(root)
    // gap + header clone before real row 2, gap before real row 4: +3
    expect(cell(root, 'a').rowSpan).toBe(8)
    expect(cell(root, 'a').getAttribute('data-base-rowspan')).toBe('5')
    // c spans real rows 0-1: ends exactly at the first insertion point
    expect(cell(root, 'c').rowSpan).toBe(2)
    expect(cell(root, 'c').hasAttribute('data-base-rowspan')).toBe(false)
    // idempotent
    syncPhantomRowspans(root)
    expect(cell(root, 'a').rowSpan).toBe(8)
    for (const tr of root.querySelectorAll('tr.page-gap, tr.page-repeat-header')) tr.remove()
    syncPhantomRowspans(root)
    expect(cell(root, 'a').rowSpan).toBe(5)
    expect(cell(root, 'a').hasAttribute('data-base-rowspan')).toBe(false)
  })

  it('leaves tables without phantom rows alone', () => {
    const root = rootOf(`
      <tr><td id="a" rowspan="2"></td><td></td></tr>
      <tr><td></td></tr>`)
    syncPhantomRowspans(root)
    expect(cell(root, 'a').rowSpan).toBe(2)
    expect(cell(root, 'a').hasAttribute('data-base-rowspan')).toBe(false)
  })
})

describe('overlay cut markers (read-only nested-table anchors)', () => {
  const rectOf = (top: number) => ({ top, height: 10 }) as DOMRect
  const wrapAt = (top: number): HTMLElement => {
    const wrap = document.createElement('div')
    wrap.getBoundingClientRect = () => rectOf(top)
    document.body.appendChild(wrap)
    return wrap
  }
  const anchorAt = (top: number): { node: Text; charOffset: number } => {
    const p = document.createElement('p')
    // empty text node: anchorTop falls back to the parent element's rect (jsdom has no Range rects)
    const node = document.createTextNode('')
    p.appendChild(node)
    p.getBoundingClientRect = () => rectOf(top)
    document.body.appendChild(p)
    return { node, charOffset: 0 }
  }

  it('creates one marker per anchor at the zoom-normalized offset', () => {
    const wrap = wrapAt(100)
    syncCutOverlays(wrap, [anchorAt(300), anchorAt(500)], 2)
    const layer = wrap.querySelector(':scope > .page-cut-overlays') as HTMLElement
    expect(layer).not.toBeNull()
    const marks = layer.querySelectorAll('.page-gap-cut.page-cut-overlay')
    expect(marks.length).toBe(2)
    expect((marks[0] as HTMLElement).style.top).toBe('100px')
    expect((marks[1] as HTMLElement).style.top).toBe('200px')
  })

  it('rebuild replaces markers; empty anchors remove the layer', () => {
    const wrap = wrapAt(0)
    syncCutOverlays(wrap, [anchorAt(50), anchorAt(60), anchorAt(70)], 1)
    syncCutOverlays(wrap, [anchorAt(80)], 1)
    const layer = wrap.querySelector('.page-cut-overlays') as HTMLElement
    expect(layer.children.length).toBe(1)
    expect((layer.children[0] as HTMLElement).style.top).toBe('80px')
    syncCutOverlays(wrap, [], 1)
    expect(wrap.querySelector('.page-cut-overlays')).toBeNull()
  })
})

describe('createLineRectsCache', () => {
  it('memoizes per element within one pass', () => {
    const rectsOf = createLineRectsCache()
    const el = document.createElement('p')
    const a = rectsOf(el, 1)
    expect(rectsOf(el, 1)).toBe(a)
    const other = document.createElement('p')
    expect(rectsOf(other, 1)).not.toBe(a)
  })
})

describe('clampCellBoxTops', () => {
  const boxAt = (pm: HTMLElement, top: number, height = 45): HTMLElement => {
    const cell = document.createElement('div')
    cell.className = 'doc-cell-boxes'
    const box = document.createElement('div')
    box.className = 'doc-textbox'
    box.getBoundingClientRect = () => ({ top, bottom: top + height, height, width: 100 }) as DOMRect
    cell.appendChild(box)
    pm.appendChild(cell)
    return box
  }

  it('pushes a box lifted above the paper top back to the edge; leaves on-page boxes alone', () => {
    const pm = document.createElement('div')
    const above = boxAt(pm, -62)
    const inside = boxAt(pm, 30)
    clampCellBoxTops(pm, 0, 1)
    expect(above.style.getPropertyValue('--page-float-dy')).toBe('62.0px')
    expect(inside.style.getPropertyValue('--page-float-dy')).toBe('')
  })

  it('is idempotent: a re-run against the already-shifted rect keeps the same dy', () => {
    const pm = document.createElement('div')
    const box = boxAt(pm, -62)
    clampCellBoxTops(pm, 0, 1)
    // after the translate the live rect reads at the paper top
    box.getBoundingClientRect = () => ({ top: 0, bottom: 45, height: 45, width: 100 }) as DOMRect
    clampCellBoxTops(pm, 0, 1)
    expect(box.style.getPropertyValue('--page-float-dy')).toBe('62.0px')
  })
})

describe('alignGapHfStrips', () => {
  const STRIP_W = 600
  /** emulates layout: stylesheet centering (left:50% of a 786px gap + translateX(-50%))
   *  until inline left/transform pin the strip */
  const stripEl = (pm: HTMLElement): HTMLElement => {
    const el = document.createElement('div')
    el.className = 'page-gap-hf'
    el.getBoundingClientRect = () => {
      const left = el.style.left ? parseFloat(el.style.left) : 393
      const tx = el.style.transform === 'none' ? 0 : -STRIP_W / 2
      return { left: left + tx, width: STRIP_W } as DOMRect
    }
    pm.appendChild(el)
    return el
  }
  const pmEl = (): HTMLElement => {
    const pm = document.createElement('div')
    pm.getBoundingClientRect = () => ({ left: 0, width: 816 }) as DOMRect
    return pm
  }

  it('pins reused stylesheet-centered strips (left:50% + translateX(-50%)) before aligning', () => {
    const pm = pmEl()
    const strip = stripEl(pm)
    alignGapHfStrips(pm, 96, 1)
    expect(strip.style.transform).toBe('none')
    expect(strip.getBoundingClientRect().left).toBeCloseTo(96, 1)
  })

  it('is idempotent across widget reuse: a second pass leaves the pinned strip alone', () => {
    const pm = pmEl()
    const strip = stripEl(pm)
    alignGapHfStrips(pm, 96, 1)
    const left = strip.style.left
    alignGapHfStrips(pm, 96, 1)
    expect(strip.style.left).toBe(left)
    expect(strip.getBoundingClientRect().left).toBeCloseTo(96, 1)
  })
})
