/**
 * Seam-fill geometry: the native selection highlight is untouched, only the
 * autospace gap column inside the selection is painted. The gap is measured
 * from the two glyph rects around the pad, so the tests stub Range rects rather
 * than a computed width.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { collectSeamRects } from '../src/renderer/editor/selection-highlight'

const rect = (left: number, top: number, width: number, height: number) => ({
  left,
  top,
  width,
  height,
})

const VIEW = { left: 0, top: 0, right: 1000, bottom: 800 }

function fixture(): { root: HTMLElement; lead: Text; pad: HTMLElement } {
  const root = document.createElement('div')
  const block = document.createElement('p')
  const lead = document.createTextNode('lead')
  const pad = document.createElement('span')
  pad.className = 'doc-autospace-pad'
  pad.textContent = 'A'
  block.append(lead, pad)
  root.append(block)
  return { root, lead, pad }
}

function rangeOver(lead: Text, pad: HTMLElement): Range {
  const range = document.createRange()
  range.setStart(lead, 0)
  const last = pad.firstChild ?? pad
  range.setEnd(last, last.textContent?.length ?? 0)
  return range
}

const ORIGINAL_RECTS = Range.prototype.getClientRects

function stubRects(map: Record<string, ReturnType<typeof rect>>): void {
  Object.defineProperty(Range.prototype, 'getClientRects', {
    configurable: true,
    value: function (this: Range) {
      const node = this.startContainer as Text
      const key = `${node.textContent}:${this.startOffset}`
      const r = map[key]
      return (r ? [r] : []) as unknown as DOMRectList
    },
  })
}

afterEach(() => {
  Object.defineProperty(Range.prototype, 'getClientRects', {
    configurable: true,
    value: ORIGINAL_RECTS,
  })
  vi.restoreAllMocks()
})

describe('collectSeamRects', () => {
  it('paints the gap between the previous glyph and the pad glyph', () => {
    const { root, lead, pad } = fixture()
    stubRects({ 'lead:3': rect(80, 10, 20, 20), 'A:0': rect(102, 10, 8, 20) })
    const out = collectSeamRects(
      root,
      [rangeOver(lead, pad)],
      VIEW,
      1,
      () => 0,
      () => rect(100, 10, 8, 20),
    )
    expect(out).toEqual([{ left: 100, top: 10, width: 2, height: 20 }])
  })

  it('skips a pad that starts the selection', () => {
    const { root, lead, pad } = fixture()
    stubRects({ 'lead:3': rect(80, 10, 20, 20), 'A:0': rect(102, 10, 8, 20) })
    const out = collectSeamRects(
      root,
      [rangeOver(lead, pad)],
      VIEW,
      5,
      () => 5,
      () => rect(100, 10, 8, 20),
    )
    expect(out).toEqual([])
  })

  it('skips gaps that are not an autospace seam', () => {
    const { root, lead, pad } = fixture()
    stubRects({ 'lead:3': rect(80, 10, 20, 20), 'A:0': rect(140, 10, 8, 20) })
    const out = collectSeamRects(
      root,
      [rangeOver(lead, pad)],
      VIEW,
      1,
      () => 0,
      () => rect(100, 10, 8, 20),
    )
    expect(out).toEqual([])
  })

  it('skips blocks outside the scroller', () => {
    const { root, lead, pad } = fixture()
    stubRects({ 'lead:3': rect(80, 10, 20, 20), 'A:0': rect(102, 10, 8, 20) })
    const out = collectSeamRects(
      root,
      [rangeOver(lead, pad)],
      { ...VIEW, top: 900, bottom: 1200 },
      1,
      () => 0,
      () => rect(100, 10, 8, 20),
    )
    expect(out).toEqual([])
  })

  it('drops boxes that fall past the scroller edge', () => {
    const { root, lead, pad } = fixture()
    stubRects({ 'lead:3': rect(990, 10, 20, 20), 'A:0': rect(1012, 10, 8, 20) })
    const out = collectSeamRects(
      root,
      [rangeOver(lead, pad)],
      VIEW,
      1,
      () => 2,
      () => rect(1000, 10, 8, 20),
    )
    expect(out).toEqual([])
  })

  it('returns nothing without a range', () => {
    const { root } = fixture()
    expect(
      collectSeamRects(
        root,
        [],
        VIEW,
        1,
        () => 0,
        () => rect(100, 10, 8, 20),
      ),
    ).toEqual([])
  })
})
