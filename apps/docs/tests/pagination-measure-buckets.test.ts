import { describe, expect, it } from 'vitest'
import { measureBlocks } from '../src/renderer/pagination'

const rectOf = (top: number, height: number, width = 100) =>
  ({
    top,
    height,
    bottom: top + height,
    left: 0,
    right: width,
    width,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect

/** a flow of top-level blocks; inner gives each block its markup */
function flowOf(specs: Array<{ top: number; height: number; inner: string; tag?: string }>) {
  const pm = document.createElement('div')
  for (const spec of specs) {
    const el = document.createElement(spec.tag ?? 'p')
    el.innerHTML = spec.inner
    el.getBoundingClientRect = () => rectOf(spec.top, spec.height)
    pm.appendChild(el)
  }
  return pm
}

describe('measureBlocks — whole-flow break scan (genoffice#526 bucketing)', () => {
  it('a page-type break in a body paragraph forces the next block to a new page', () => {
    const pm = flowOf([
      { top: 0, height: 20, inner: 'one' },
      { top: 20, height: 20, inner: 'two<br class="doc-page-br">' },
      { top: 40, height: 20, inner: 'after' },
    ])
    const { blocks } = measureBlocks(pm, 0, 1)
    // the trailing break pushes the next block (the slicer consumes breakAfter;
    // measureBlocks itself sets no breakBefore there)
    expect(blocks[1].breakAfter).toBe(true)
    expect(blocks[2].breakBefore).toBeUndefined()
  })

  it('a page-type break inside a table cell never breaks the body flow', () => {
    const pm = flowOf([
      { top: 0, height: 20, inner: 'one' },
      {
        top: 20,
        height: 40,
        inner: '<table><tr><td>cell<br class="doc-page-br">text</td></tr></table>',
        tag: 'div',
      },
      { top: 60, height: 20, inner: 'after' },
    ])
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[1].breakAfter).toBeUndefined()
    expect(blocks[2].breakBefore).toBeUndefined()
    // the table still marks the block non-reflowable
    expect(blocks[1].fixedWidthPx).toBeDefined()
    expect(blocks[0].fixedWidthPx).toBeUndefined()
  })

  it('a break inside a textbox lays out the box, not the body flow', () => {
    const pm = flowOf([
      { top: 0, height: 20, inner: 'one' },
      { top: 20, height: 30, inner: '<div class="doc-textbox">box<br class="doc-page-br"></div>' },
      { top: 50, height: 20, inner: 'after' },
    ])
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[1].breakAfter).toBeUndefined()
    expect(blocks[2].breakBefore).toBeUndefined()
  })

  it('in-block inline gap bands subtract from the block height', () => {
    const pm = flowOf([
      {
        top: 0,
        height: 60,
        inner: 'before<div class="page-gap-inline" style="height:18px"></div>after',
      },
      { top: 60, height: 20, inner: 'next' },
    ])
    // jsdom lays out nothing: give the band the rect the browser would
    for (const g of pm.querySelectorAll('.page-gap-inline'))
      (g as HTMLElement).getBoundingClientRect = () => rectOf(40, 18)
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[0].height).toBe(42)
  })

  it('an image-bearing paragraph is not an empty paragraph', () => {
    const pm = flowOf([
      { top: 0, height: 20, inner: '<img src="x.png">' },
      { top: 20, height: 20, inner: '' },
    ])
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[0].emptyPara).toBeUndefined()
    expect(blocks[1].emptyPara).toBe(true)
  })

  it('a leading break found through nested markup leads its block to the next page', () => {
    const pm = flowOf([
      {
        top: 0,
        height: 50,
        inner: '<div><span><em><br class="doc-page-br">c</em></span></div>',
      },
      { top: 50, height: 20, inner: 'after' },
    ])
    const { blocks } = measureBlocks(pm, 0, 1)
    // no text before the break, text after: the break leads (Word keeps the
    // break line here and starts the text on the next page). The block has
    // text, so it is not a break-only paragraph (no dedicated line height).
    expect(blocks[0].breakBefore).toBe(true)
    expect(blocks[0].breakOnlyLineH).toBeUndefined()
  })

  it('bucketing agrees with the old per-block scan on a mixed flow', () => {
    const pm = flowOf([
      { top: 0, height: 20, inner: 'plain' },
      { top: 20, height: 20, inner: 'lead<br class="doc-page-br">' },
      {
        top: 40,
        height: 30,
        inner: '<table><tr><td>x</td></tr></table>',
        tag: 'div',
      },
      {
        top: 70,
        height: 40,
        inner: 'a<div class="page-gap-inline" style="height:10px"></div>b<br class="doc-page-br">',
      },
      { top: 110, height: 20, inner: '<img src="y.png">' },
    ])
    for (const g of pm.querySelectorAll('.page-gap-inline'))
      (g as HTMLElement).getBoundingClientRect = () => rectOf(80, 10)
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks).toHaveLength(5)
    expect(blocks[1].breakAfter).toBe(true)
    expect(blocks[2].fixedWidthPx).toBeDefined()
    expect(blocks[3].height).toBe(30)
    expect(blocks[3].breakAfter).toBe(true)
    expect(blocks[4].emptyPara).toBeUndefined()
  })
})
