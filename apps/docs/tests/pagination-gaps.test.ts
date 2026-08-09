import { describe, expect, it } from 'vitest'
import { GAP_BAND, makeGapEl } from '../src/renderer/editor/pagination-gaps'
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
