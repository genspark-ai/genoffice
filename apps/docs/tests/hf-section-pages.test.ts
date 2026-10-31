import { describe, expect, it } from 'vitest'
import {
  PAGE_MARK,
  SECTION_PAGES_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
} from '@genoffice/docx-engine'
import { makeGapHfEl } from '../src/renderer/editor/hf-dom'
import { hfDocJson, hfValueFromDoc } from '../src/renderer/editor/hf-editor'
import { applyHfText, hfEditText } from '../src/renderer/editor/hf-text'
import { sectionPageCounts, type PageSlice } from '../src/renderer/pagination'

/** SECTIONPAGES (#2331): the strip shows the page count of its own section */

const FOOTER: HeaderFooter = {
  text: `- ${PAGE_MARK} / ${SECTION_PAGES_MARK} -`,
  paras: [{ align: 'center', runs: [{ text: `- ${PAGE_MARK} / ${SECTION_PAGES_MARK} -` }] }],
}

const slice = (start: number, end: number, section: number): PageSlice => ({ start, end, section })

describe('SECTIONPAGES in header/footer strips', () => {
  it('a gap strip fills the mark with the section page count', () => {
    const el = makeGapHfEl({
      kind: 'footer',
      value: FOOTER,
      pageNo: 9,
      pageTotal: 12,
      sectionPages: 8,
    })
    expect(el.textContent).toBe('- 9 / 8 -')
  })

  it('without a section count the mark falls back to the total, never an empty string', () => {
    const el = makeGapHfEl({ kind: 'footer', value: FOOTER, pageNo: 2, pageTotal: 5 })
    expect(el.textContent).toBe('- 2 / 5 -')
  })

  it('counts pages per run of slices sharing a section', () => {
    const slices = [
      slice(0, 10, 0),
      slice(10, 20, 0),
      slice(20, 30, 0),
      slice(30, 40, 1),
      slice(40, 50, 1),
    ]
    expect(sectionPageCounts(slices)).toEqual([3, 3, 3, 2, 2])
    expect(sectionPageCounts([])).toEqual([])
  })

  it('the strip editor keeps the field as an atom and writes the mark back', () => {
    const doc = hfDocJson(FOOTER)
    const fields = (doc.content?.[0].content ?? []).filter((n) => n.type === 'hfField')
    expect(fields.map((n) => n.attrs?.kind)).toEqual(['PAGE', 'SECTIONPAGES'])
    const back = hfValueFromDoc(doc, FOOTER)
    expect(back.paras![0].runs.map((r) => r.text).join('')).toBe(FOOTER.text)
  })

  it('the plain-text view shows {SECTIONPAGES} and maps it back', () => {
    const text = hfEditText(FOOTER)
    expect(text).toBe('- {PAGE} / {SECTIONPAGES} -')
    const next = applyHfText(FOOTER, '{PAGE} of {SECTIONPAGES} ({NUMPAGES})')
    expect(next.paras?.[0].runs.map((r) => r.text).join('')).toBe(
      `${PAGE_MARK} of ${SECTION_PAGES_MARK} (${TOTAL_PAGES_MARK})`,
    )
  })
})
