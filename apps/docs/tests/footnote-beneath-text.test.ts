import { describe, expect, it } from 'vitest'
import { appendEndnotesBlock, noteAreaPlacement, pageTextEnds } from '../src/renderer/pagination'
import type { BlockBox, PageNoteItem, PageSlice } from '../src/renderer/pagination-types'

const el = (marginBottomPx = 0): HTMLElement => {
  const d = document.createElement('div')
  d.style.marginBottom = `${marginBottomPx}px`
  document.body.appendChild(d)
  return d
}

const slice = (start: number, end: number): PageSlice => ({ start, end, section: 0 })

describe('footnotes beneath text (w:footnotePr w:pos="beneathText")', () => {
  it('page text end excludes the footnote reservation and clips a split block at the page end', () => {
    const blocks: BlockBox[] = [
      // referencing paragraph: 100px of text + 40px reserved for its note
      { top: 0, height: 140, footnoteExtraPx: 40, el: el(), docxIndex: 0 },
      { top: 100, height: 300, el: el(), docxIndex: 1 },
      { top: 400, height: 50, el: el(), docxIndex: 2 },
    ]
    const slices = [slice(0, 250), slice(250, 450)]
    // page 1 ends inside the second block; page 2 ends at the last block's bottom
    expect(pageTextEnds(blocks, slices)).toEqual([250, 450])
  })

  it('the flow-final block contributes its space after (Word lays the separator below it)', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 100, footnoteExtraPx: 20, el: el(), docxIndex: 0 },
      { top: 80, height: 40, el: el(13), docxIndex: 1 },
    ]
    expect(pageTextEnds(blocks, [slice(0, 500)])).toEqual([133])
  })

  it('virtual blocks (endnote area, float spill) are not text', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 60, el: el(), docxIndex: 0 },
      { top: 60, height: 200, isEndnotes: true },
      { top: 260, height: 300 },
    ]
    expect(pageTextEnds(blocks, [slice(0, 700)])).toEqual([60])
  })
})

describe('page note area placement', () => {
  const geom = { pageH: 1000, mTop: 100, mBottom: 100, headerH: 0, vOffset: 0 }

  it('bottom-anchored by default; beneath the text end when the section asks for it', () => {
    expect(noteAreaPlacement(60, undefined, slice(0, 800), geom)).toEqual({
      top: null,
      endnoteShift: 0,
    })
    expect(noteAreaPlacement(60, 300, slice(0, 800), geom)).toEqual({ top: 400, endnoteShift: 60 })
    // an over-tall area clamps to the content bottom; endnotes follow its real bottom
    expect(noteAreaPlacement(60, 790, slice(0, 800), geom)).toEqual({ top: 840, endnoteShift: 10 })
  })

  it('moves with the page vAlign shift like the body text', () => {
    expect(noteAreaPlacement(60, 300, slice(0, 800), { ...geom, vOffset: 50 })).toEqual({
      top: 450,
      endnoteShift: 60,
    })
  })

  it('endnotes on a beneath-text page stack below the footnote area, not over it', () => {
    const { top, endnoteShift } = noteAreaPlacement(60, 300, slice(0, 800), geom)
    // endnote rows start at the flow end (300) like the footnote area; the shift keeps them apart
    const endnoteTop = geom.mTop + (300 - 0) + endnoteShift
    expect(endnoteTop).toBe(top! + 60)
  })

  it('no footnotes or regioned pages leave the endnotes at their flow position', () => {
    expect(noteAreaPlacement(0, 300, slice(0, 800), geom)).toEqual({ top: null, endnoteShift: 0 })
    expect(noteAreaPlacement(60, 300, { ...slice(0, 800), regions: [] }, geom)).toEqual({
      top: null,
      endnoteShift: 0,
    })
  })
})

describe('endnote area anchor', () => {
  it('starts below the last body paragraph including its space after', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 100, el: el(13), docxIndex: 0 }]
    const items: PageNoteItem[] = [{ no: 1, id: '1', text: 'n', height: 20 }]
    const out = appendEndnotesBlock(blocks, 100, items, 16)
    expect(out).toEqual({ top: 113, totalHeight: 149 })
    expect(blocks[1]).toMatchObject({ top: 113, height: 36, isEndnotes: true })
  })

  it('a folded inter-block margin is not counted twice', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 113, spaceAfterPx: 13, el: el(13), docxIndex: 0 }]
    const out = appendEndnotesBlock(blocks, 113, [{ no: 1, id: '1', text: 'n', height: 20 }], 16)
    expect(out?.top).toBe(113)
  })
})
