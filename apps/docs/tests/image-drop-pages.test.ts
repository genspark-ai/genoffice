/**
 * Where a dragged picture lands among pages read off the page gaps: the page
 * its top-left corner lands on (a top in the gray band belongs to the next
 * sheet), pulled inside that paper, anchored to a paragraph of that page low
 * enough for the picture to fit below it.
 */
import { describe, expect, it } from 'vitest'
import {
  imageDropPages,
  placeImageDrop,
  type ImageParagraphAnchor,
} from '../src/renderer/editor/extensions'
import { GAP_BAND } from '../src/renderer/editor/pagination-gaps'

const rect = (top: number, bottom: number, left = 100, right = 700) =>
  ({ top, bottom, left, right, height: bottom - top, width: right - left }) as DOMRect

/** a page gap at client [top, bottom) closing a page with the given text area */
function gapEl(top: number, bottom: number, data: { bodyHeight?: number; bodyPad?: number }) {
  const gap = document.createElement('div')
  gap.className = 'page-gap'
  gap.style.setProperty('--gap-mt', '96px')
  if (data.bodyHeight != null) gap.dataset.bodyHeight = String(data.bodyHeight)
  if (data.bodyPad != null) gap.dataset.bodyPad = String(data.bodyPad)
  gap.getBoundingClientRect = () => rect(top, bottom)
  return gap
}

/**
 * Three pages: page 1's text ends 40 px of unused body below the first gap's
 * top (1000); page 2's body starts at 1300 and is 1000 px tall; page 3 has no
 * gap below it.
 */
function pages() {
  const pm = document.createElement('div')
  pm.append(gapEl(1000, 1300, { bodyPad: 40 }), gapEl(2300, 2600, { bodyHeight: 1000 }))
  return imageDropPages(pm, 1)
}

const anchor = (top: number, type = 'docParagraph'): ImageParagraphAnchor =>
  ({
    pos: top,
    node: { type: { name: type }, attrs: {} },
    dom: document.createElement('p'),
    rect: rect(top, top + 40),
  }) as unknown as ImageParagraphAnchor

describe('imageDropPages', () => {
  it('reads paper and text area of each page off the gaps', () => {
    const [p1, p2, p3] = pages()
    // the next paper starts --gap-mt above the gap's bottom; GAP_BAND of gray before it
    expect(p1).toMatchObject({ bodyBottom: 1040, paperBottom: 1300 - 96 - GAP_BAND })
    expect(p2).toEqual({
      paperTop: 1204,
      bodyTop: 1300,
      bodyBottom: 2300,
      paperBottom: 2600 - 96 - GAP_BAND,
    })
    // the last page has no gap below it: it takes the extents of the one before
    expect(p3).toEqual({ paperTop: 2504, bodyTop: 2600, bodyBottom: 3600, paperBottom: 3776 })
  })
})

describe('placeImageDrop', () => {
  const anchors = [anchor(400), anchor(900), anchor(1300), anchor(1800), anchor(2200), anchor(2600)]

  it('a top in the gray band between two sheets belongs to the next one', () => {
    const band = 1300 - 96 - GAP_BAND + 10
    const placed = placeImageDrop(anchors, 300, band, 100, pages(), 1)
    // pulled inside page 2's paper, anchored to page 2's first paragraph
    expect(placed.top).toBe(1204)
    expect(placed.anchor?.rect.top).toBe(1300)
  })

  it('a picture hanging off the bottom of its page is pulled up inside that paper', () => {
    const placed = placeImageDrop(anchors, 300, 2350, 200, pages(), 1)
    expect(placed.top).toBe(2476 - 200)
    // the paragraph at 2200 leaves no room below it: the page would turn
    // before it and carry the picture along, so the lowest one that fits
    expect(placed.anchor?.rect.top).toBe(1800)
  })

  it('a heading needs room for the next paragraph line too', () => {
    const heading = [anchor(1300), anchor(2000, 'docHeading')]
    // 2000 + 260 fits under a paragraph (2276) but not under a heading (2252)
    expect(placeImageDrop(heading, 300, 2050, 260, pages(), 1).anchor?.rect.top).toBe(1300)
    const para = [anchor(1300), anchor(2000)]
    expect(placeImageDrop(para, 300, 2050, 260, pages(), 1).anchor?.rect.top).toBe(2000)
  })

  it('a drop inside the text area keeps its point and the paragraph under it', () => {
    const placed = placeImageDrop(anchors, 300, 1820, 100, pages(), 1)
    expect(placed).toMatchObject({ top: 1820 })
    expect(placed.anchor?.rect.top).toBe(1800)
  })
})
