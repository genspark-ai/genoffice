import { describe, expect, it } from 'vitest'
import { GAP_BAND, makeGapEl, syncPhantomRowspans } from '../src/renderer/editor/pagination-gaps'
import { singleCutCell } from '../src/renderer/pagination'

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

  it('multi-cell / missing rows keep the zero-height cut marker', () => {
    expect(singleCutCell(rowOf(2))).toBeNull()
    expect(singleCutCell(null)).toBeNull()
    expect(makeGapEl(m, 'cut').className).toBe('page-gap-cut')
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
