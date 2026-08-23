import { describe, expect, it } from 'vitest'
import { syncAnchorBands, syncFloatShifts } from '../src/renderer/editor/pagination-gaps'

const rectOf = (top: number, height: number) =>
  ({ top, bottom: top + height, height, width: 100, left: 0, right: 100 }) as DOMRect

/** floating anchor wrapper with a line strut and wrapTopAndBottom band data */
const wrapperOf = (bands: string, lineH = 20): HTMLElement => {
  const el = document.createElement('div')
  el.className = 'doc-protected doc-protected-textboxes doc-protected-floating'
  if (bands) {
    const bottoms = bands.split(' ').map((s) => parseInt(s.split(':')[1], 10))
    el.dataset.band = String(Math.max(...bottoms))
    el.dataset.bands = bands
    el.style.minHeight = `${el.dataset.band}px`
  }
  const strut = document.createElement('div')
  strut.className = 'doc-anchor-strut'
  strut.getBoundingClientRect = () => rectOf(0, lineH)
  el.appendChild(strut)
  return el
}

describe('syncAnchorBands', () => {
  it('collapses a run onto anchor lines + the band union (photo wall)', () => {
    const pm = document.createElement('div')
    // side-by-side photos overlap vertically: coverage counts their union, so
    // the second anchor's line still slips in above the first photo row
    const w1 = wrapperOf('32:150 33:152 477:508')
    const w2 = wrapperOf('0:126 353:448')
    pm.append(w1, w2)
    syncAnchorBands(pm, 1)
    // w2's line [20,40) only grazes w1's band at 32 (8px < half a line): it
    // stays one line below w1, and w2 extends to the union bottom 508
    expect(w1.style.minHeight).toBe('20px')
    expect(w2.style.minHeight).toBe('488px')
  })

  it('pushes an anchor line below bands that cover it (Word band exclusion)', () => {
    const pm = document.createElement('div')
    const w1 = wrapperOf('10:327')
    const w2 = wrapperOf('0:100')
    pm.append(w1, w2)
    syncAnchorBands(pm, 1)
    // w2's line [20,40) is fully inside w1's band: its origin moves to 327
    expect(w1.style.minHeight).toBe('327px')
    expect(w2.style.minHeight).toBe('100px')
  })

  it('restores a single wrapper to its own band after a run dissolves', () => {
    const pm = document.createElement('div')
    const w1 = wrapperOf('0:300')
    w1.dataset.bandAdj = '0'
    w1.style.minHeight = '0px'
    pm.append(w1)
    syncAnchorBands(pm, 1)
    expect(w1.style.minHeight).toBe('300px')
    expect(w1.dataset.bandAdj).toBeUndefined()
  })
})

describe('syncFloatShifts with in-table gaps', () => {
  it('counts gap rows inside tables when shifting floats below them', () => {
    const pm = document.createElement('div')
    pm.getBoundingClientRect = () => rectOf(0, 1000)
    const table = document.createElement('table')
    const gapRow = document.createElement('tr')
    gapRow.className = 'page-gap page-gap-inline page-gap-table'
    gapRow.getBoundingClientRect = () => rectOf(200, 100)
    table.appendChild(gapRow)
    const wrap = document.createElement('div')
    wrap.className = 'doc-protected doc-protected-floating'
    const box = document.createElement('div')
    box.className = 'doc-textbox'
    box.getBoundingClientRect = () => rectOf(250, 50)
    wrap.appendChild(box)
    pm.append(table, wrap)
    // virtual top 250 lies below the gap cut at v=200: the box shifts down by
    // the in-table gap height so it does not resolve one page too high
    syncFloatShifts(pm, [{ el: box, top: 250 }], 0, 1)
    expect(box.style.getPropertyValue('--page-float-dy')).toBe('100.0px')
  })
})
