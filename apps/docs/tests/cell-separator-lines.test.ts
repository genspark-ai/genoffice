/**
 * An empty paragraph holding an inline widget (a bookmark, an editable-region
 * mark) carries ProseMirror's empty `img.ProseMirror-separator`. In a table
 * cell it was read as a picture, so the paragraph lost its own line: blank
 * lines between a cell's text and a picture below it became one gap, cut in
 * its middle, and a blank line Word keeps at the page bottom went over.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { fillLineBoxes, measureBlocks } from '../src/renderer/pagination'

const rectOf = (top: number, height: number, width = 400) =>
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

const at = (el: Element, top: number, height: number, width = 400) => {
  el.getBoundingClientRect = () => rectOf(top, height, width)
  return el
}

let restore: () => void = () => {}
beforeEach(() => {
  // jsdom has no Range.getClientRects: a text node reports its paragraph's box
  const range = Range.prototype as unknown as { getClientRects?: () => DOMRect[] }
  const orig = range.getClientRects
  range.getClientRects = function (this: Range) {
    const el = this.startContainer.parentElement
    return el && this.startContainer.textContent?.trim() ? [el.getBoundingClientRect()] : []
  }
  restore = () => (range.getClientRects = orig)
})
afterEach(() => restore())

/** a one-cell row: a text line, `blanks` empty paragraphs, then a picture */
function cellWithBlanks(withSeparators: boolean, blanks = 2) {
  const pm = document.createElement('div')
  const table = document.createElement('table')
  table.className = 'doc-table'
  table.style.borderCollapse = 'collapse'
  const tbody = document.createElement('tbody')
  const tr = document.createElement('tr')
  const td = document.createElement('td')
  const text = document.createElement('p')
  text.textContent = 'illustrated in Figure 2.'
  at(text, 0, 20)
  td.appendChild(text)
  for (let i = 0; i < blanks; i++) {
    const p = document.createElement('p')
    p.className = 'doc-p-empty'
    p.innerHTML = withSeparators
      ? '<span class="perm-region-mark"></span><img class="ProseMirror-separator" alt=""><br class="ProseMirror-trailingBreak">'
      : '<br class="ProseMirror-trailingBreak">'
    at(p, 20 + i * 20, 20)
    const sep = p.querySelector('img')
    if (sep) at(sep, 20 + i * 20, 0, 0)
    td.appendChild(p)
  }
  const pic = document.createElement('p')
  const img = document.createElement('img')
  pic.appendChild(img)
  const picTop = 20 + blanks * 20
  at(pic, picTop, 300)
  at(img, picTop, 300)
  td.appendChild(pic)
  const h = picTop + 300
  for (const el of [td, tr, table]) at(el, 0, h)
  tr.appendChild(td)
  tbody.appendChild(tr)
  table.appendChild(tbody)
  pm.appendChild(table)
  document.body.replaceChildren(pm)
  const { blocks } = measureBlocks(pm, 0, 1)
  fillLineBoxes(
    blocks,
    [{ contentHeight: 800, forceBreak: false }],
    1,
    undefined,
    undefined,
    new Set(blocks),
  )
  return blocks[0]!.tableRows![0]!
}

describe('empty paragraphs in a table cell', () => {
  it('stay lines of their own when they carry an editor separator', () => {
    const plain = cellWithBlanks(false)
    const withSeparators = cellWithBlanks(true)
    // a cut at each blank line's edge, the last one right above the picture
    expect(plain.cutYs).toEqual(expect.arrayContaining([20, 40, 60]))
    expect(withSeparators.cutYs).toEqual(plain.cutYs)
    expect(withSeparators.contentBottom).toBe(plain.contentBottom)
  })
})
