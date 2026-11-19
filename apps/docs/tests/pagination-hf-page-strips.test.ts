import { describe, expect, it } from 'vitest'
import type { HeaderFooter, SectionInfo } from '@genoffice/docx-engine'
import { EMPTY_HF_VARIANTS } from '../src/renderer/doc-state'
import { resolveHf, withHfEdit, type HfSectionState } from '../src/renderer/hf-sections'
import { hfVariantOf, pageHfStrips, type HfSet } from '../src/renderer/pagination-hf'

const hf = (text: string): HeaderFooter => ({ text, paras: [{ runs: [{ text }] }] })

const loneSection: SectionInfo = {
  settings: {} as SectionInfo['settings'],
  startType: 'nextPage',
  firstBlockIndex: 0,
  lastBlockIndex: 5,
  sectPrXml: '',
  titlePg: false,
  headerRefs: {},
  footerRefs: {},
}

/** what the preview receives as `hf`: the parsed strips, never the pending edits */
const parsed: HfSet = {
  header: hf('parsed header'),
  footer: null,
  headerFirst: hf('parsed first'),
  footerFirst: null,
  headerEven: null,
  footerEven: null,
  titlePg: false,
  evenOddHf: false,
  images: { header: [{ dataUrl: 'logo' }] },
}

function state(sections: SectionInfo[]): HfSectionState {
  return {
    sections,
    edits: {},
    links: {},
    docStrips: { header: parsed.header, footer: parsed.footer, variants: { ...EMPTY_HF_VARIANTS } },
    parsed: { headerImages: parsed.images?.header },
  }
}

const resolverOf = (s: HfSectionState, variant: 'default' | 'first' | 'even') => {
  return (kind: 'header' | 'footer') => resolveHf(s, 0, kind, variant)
}

describe('pageHfStrips (one-section preview)', () => {
  it('shows the pending strip edit, not the parsed header', () => {
    const s = state([loneSection])
    const edited = { ...s, ...withHfEdit(s, 0, 'header', 'default', hf('typed header'))! }
    const page = pageHfStrips('default', resolverOf(edited, 'default'), parsed)
    expect(page.header?.text).toBe('typed header')
    expect(page.headerImages).toEqual([{ dataUrl: 'logo' }])
  })

  it('a footer added to a document that had none is drawn', () => {
    const s = state([loneSection])
    const edited = { ...s, ...withHfEdit(s, 0, 'footer', 'default', hf('page footer'))! }
    const page = pageHfStrips('default', resolverOf(edited, 'default'), parsed)
    expect(page.footer?.text).toBe('page footer')
  })

  it('without pending edits the resolver and the parsed set agree', () => {
    const s = state([loneSection])
    expect(pageHfStrips('default', resolverOf(s, 'default'), parsed)).toEqual(
      pageHfStrips('default', undefined, parsed),
    )
  })

  it('the fallback picks the variant slot of the parsed set', () => {
    const variant = hfVariantOf(true, true, false, 1)
    expect(variant).toBe('first')
    expect(pageHfStrips(variant, undefined, parsed).header?.text).toBe('parsed first')
    expect(pageHfStrips('even', undefined, parsed).header).toBeNull()
  })
})
