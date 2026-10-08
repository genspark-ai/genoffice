import { describe, expect, it } from 'vitest'
import { PAGE_MARK, type HfParagraph } from '@genoffice/docx-engine'
import {
  hfTabSegments,
  hfWithoutPageMarks,
  makeGapHfEl,
  type HfStripGeom,
} from '../src/renderer/editor/hf-dom'
import { hfDocJson, hfValueFromDoc } from '../src/renderer/editor/hf-editor'
import { applyHfText } from '../src/renderer/editor/hf-text'

/**
 * A header paragraph "text <tab> inline picture" with a left stop whose segment
 * ends exactly at the column edge: Word draws the picture at the stop on the
 * text line, its bottom on the shared baseline.
 */

const geom: HfStripGeom = {
  pageW: 794,
  pageH: 1123,
  marginLeft: 84,
  marginRight: 84,
  marginTop: 96,
  marginBottom: 96,
  headerStripTop: 47,
  footerDist: 47,
}

const picture = { dataUrl: 'data:image/png;base64,AA==', widthPx: 154, heightPx: 58, xml: '' }
const para: HfParagraph = {
  tabStops: [{ val: 'left', pos: 7074 }],
  runs: [{ text: 'Brand | Template' }, { text: '\t' }, { text: '', image: picture }],
}

describe('inline picture after a header tab', () => {
  it('keeps the picture run in the tab segment at its stop', () => {
    const layout = hfTabSegments(para)!
    expect(layout.lead.map((r) => r.text)).toEqual(['Brand | Template'])
    expect(layout.segments).toHaveLength(1)
    expect(layout.segments[0].left).toEqual({ px: 7074 / 15 })
    expect(layout.segments[0].runs[0].image).toBe(picture)
  })

  it('renders the picture inside the positioned segment with a baseline strut', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: { text: '', paras: [para] },
      pageNo: 1,
      pageTotal: 1,
      geom,
    })
    expect(el.querySelector('.page-hf-images')).toBeNull()
    const seg = el.querySelector<HTMLElement>('.page-hf-tabseg')!
    expect(seg.classList.contains('page-hf-tabseg-img')).toBe(true)
    expect(seg.style.left).toBe(`${(7074 / 15).toFixed(1)}px`)
    const img = seg.querySelector<HTMLImageElement>('img.page-hf-run-img')!
    expect(img.style.width).toBe('154px')
    const strut = el.querySelector<HTMLElement>('.page-hf-tab-strut')!
    expect(strut.style.height).toBe('58px')
    // the picture and the text share one line box
    expect(el.querySelectorAll('.page-hf-para')).toHaveLength(1)
  })
})

describe('pictures survive header text edits', () => {
  it('applyHfText keeps a leading picture first and a trailing one last', () => {
    const lead: HfParagraph = { runs: [{ text: '', image: picture }, { text: 'Title' }] }
    const trail: HfParagraph = { runs: [{ text: 'Title\t' }, { text: '', image: picture }] }
    const edited = applyHfText({ text: 'Title', paras: [lead, trail] }, 'A\nB')
    expect(edited.paras![0].runs.map((r) => [r.text, !!r.image])).toEqual([
      ['', true],
      ['A', false],
    ])
    expect(edited.paras![1].runs.map((r) => [r.text, !!r.image])).toEqual([
      ['B', false],
      ['', true],
    ])
  })

  it('hfDocJson leaves pictures out of the editor and hfValueFromDoc restores them', () => {
    const base = {
      text: 'Title',
      paras: [{ runs: [{ text: '', image: picture }, { text: 'Title' }] }],
    }
    const doc = hfDocJson(base)
    expect(JSON.stringify(doc)).not.toContain('image')
    const back = hfValueFromDoc(doc, base)
    expect(back.paras![0].runs.map((r) => [r.text, !!r.image])).toEqual([
      ['', true],
      ['Title', false],
    ])
  })

  it('hfWithoutPageMarks keeps a paragraph that is only a picture once the field is gone', () => {
    const out = hfWithoutPageMarks({
      text: PAGE_MARK,
      pageNumber: true,
      paras: [{ runs: [{ text: PAGE_MARK }, { text: '', image: picture }] }],
    })
    expect(out.paras).toHaveLength(1)
    expect(out.paras![0].runs.map((r) => [r.text, !!r.image])).toEqual([['', true]])
  })
})
