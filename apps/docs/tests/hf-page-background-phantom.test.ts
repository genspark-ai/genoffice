import { describe, expect, it } from 'vitest'
import type { HeaderFooter, HfPartInfo, SectionInfo, StyleInfo } from '@genoffice/docx-engine'
import { hfLayoutResolved, hfPhantomSpec, hfWithPhantom } from '../src/renderer/hf-phantom'
import { resolveHf, type HfKind, type HfSectionState } from '../src/renderer/hf-sections'
import { estimateHfHeight, lineHeightFactor } from '../src/renderer/line-metrics'
import type { HfView } from '../src/renderer/doc-state'

// Normal = Georgia 13 / line 300 / after 120; Header basedOn Normal with after 0 / line 240
const styles = new Map<string, StyleInfo>([
  [
    'Normal',
    {
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: {
        sizeHalfPoints: 26,
        fontAscii: 'Georgia',
        font: 'Georgia',
        lineRule: 'auto',
        lineRawTwips: 300,
        lineSpacing: 1.25,
        spaceAfterTwips: 120,
      },
    },
  ],
  [
    'Header',
    {
      styleId: 'Header',
      name: 'header',
      type: 'paragraph',
      basedOn: 'Normal',
      display: {
        sizeHalfPoints: 26,
        fontAscii: 'Georgia',
        font: 'Georgia',
        lineRule: 'auto',
        lineRawTwips: 240,
        lineSpacing: 1,
        spaceAfterTwips: 0,
      },
    },
  ],
])
const docDefaults = { sizeHalfPoints: 20, spaceAfterTwips: 200, lineSpacing: 1.15 }
const spec = hfPhantomSpec({ styles, docDefaults })

// the one-line Header-style header: Georgia 8.5 italic, line 240, after 0
const headerPara = {
  align: 'right' as const,
  lineRule: 'auto' as const,
  lineRawTwips: 240,
  lineSpacing: 1,
  spaceAfter: 0,
  runs: [{ text: 'Report · 2026', font: 'Georgia', sizeHalfPoints: 17, italic: true }],
}
const header: HeaderFooter = { text: 'Report', paras: [headerPara] }
const part = (text: string): HfPartInfo => ({ text, hasPageNumber: false, paras: [headerPara] })

function section(
  lastBlockIndex: number,
  headerRefs: SectionInfo['headerRefs'] = {},
  footerRefs: SectionInfo['footerRefs'] = {},
): SectionInfo {
  return {
    settings: {} as SectionInfo['settings'],
    startType: 'nextPage',
    firstBlockIndex: 0,
    lastBlockIndex,
    sectPrXml: '',
    titlePg: false,
    headerRefs,
    footerRefs,
  }
}

const state = (sections: SectionInfo[]): HfSectionState => ({
  sections,
  edits: {},
  links: {},
  parsed: { hfParts: { h1: part('one'), h2: part('two'), f1: part('foot') } },
})

const layout = (st: HfSectionState, si: number, kind: HfKind, variant: HfView = 'default') =>
  hfLayoutResolved(kind, resolveHf(st, si, kind, variant), spec)
const parasOf = (v: HeaderFooter | null) => v?.paras ?? []

describe('page background header phantom', () => {
  it('builds the blank Normal paragraph from the default paragraph style', () => {
    expect(spec.normal).toMatchObject({
      runs: [],
      align: 'left',
      emptyRunSizeHalfPoints: 26,
      emptyRunFontFamily: 'Georgia',
      lineRule: 'auto',
      lineRawTwips: 300,
      lineSpacing: 1.25,
      spaceBefore: 0,
      spaceAfter: 120,
    })
    expect(spec.headerLine).toMatchObject({ lineSpacing: 1, spaceAfter: 0 })
  })

  it('falls back to docDefaults when Normal declares no spacing', () => {
    const bare = new Map<string, StyleInfo>([
      ['Normal', { styleId: 'Normal', name: 'Normal', type: 'paragraph', isDefault: true }],
    ])
    expect(hfPhantomSpec({ styles: bare, docDefaults }).normal).toMatchObject({
      emptyRunSizeHalfPoints: 20,
      lineSpacing: 1.15,
      spaceAfter: 200,
    })
  })

  it('appends one phantom after the header content, never to the source value', () => {
    const out = hfWithPhantom(header, spec)
    expect(parasOf(out)).toHaveLength(2)
    expect(parasOf(out)[1]).toEqual(spec.normal)
    expect(header.paras).toHaveLength(1)
  })

  it('merges the last header after with Normal before by max', () => {
    const before = { ...spec, normal: { ...spec.normal, spaceBefore: 200 } }
    expect(parasOf(hfWithPhantom(header, before))[1].spaceBefore).toBe(200)
    const after = { text: '', paras: [{ ...headerPara, spaceAfter: 140 }] }
    expect(parasOf(hfWithPhantom(after, before))[1].spaceBefore).toBe(60)
  })

  it('section 1 without a header part gets an empty Header line and the phantom', () => {
    const out = hfWithPhantom(null, spec)
    expect(parasOf(out)).toEqual([spec.headerLine, spec.normal])
  })

  it('every header variant of section 1, and a section inheriting it', () => {
    const st = state([
      section(5, { default: 'h1', first: 'h2', even: 'h1' }, { default: 'f1' }),
      section(9),
    ])
    for (const variant of ['default', 'first', 'even'] as const) {
      expect(parasOf(layout(st, 0, 'header', variant).value)).toHaveLength(2)
    }
    const inherited = layout(st, 1, 'header')
    expect(inherited.owner).toBe(0)
    expect(parasOf(inherited.value)).toHaveLength(2)
  })

  it('not for a section with its own header part, never for footers', () => {
    const st = state([
      section(5, { default: 'h1' }, { default: 'f1' }),
      section(9, { default: 'h2' }),
    ])
    const own = layout(st, 1, 'header')
    expect(own.owner).toBe(1)
    expect(parasOf(own.value)).toHaveLength(1)
    expect(parasOf(layout(st, 0, 'footer').value)).toHaveLength(1)
    expect(parasOf(layout(st, 1, 'footer').value)).toHaveLength(1)
  })

  it('a header found nowhere inherits the part-less section 1 story', () => {
    const st = state([section(5, {}, { default: 'f1' }), section(9)])
    expect(parasOf(layout(st, 1, 'header').value)).toEqual([spec.headerLine, spec.normal])
  })

  it('nothing without a page background', () => {
    const st = state([section(5, { default: 'h1' })])
    const plain = hfLayoutResolved('header', resolveHf(st, 0, 'header', 'default'), null)
    expect(parasOf(plain.value)).toHaveLength(1)
  })

  it('reserves Normal sz x font factor x line multiple + Normal after', () => {
    const w = ((12240 - 2 * 1440) / 1440) * 96
    const plain = estimateHfHeight(header, w)
    const withPhantom = estimateHfHeight(hfWithPhantom(header, spec), w)
    const phantomPt = 13 * lineHeightFactor('Georgia') * 1.25 + 6
    expect(((withPhantom - plain) * 72) / 96).toBeCloseTo(phantomPt, 1)
    // Georgia 13 x 1.25 + 6pt after = 24.5pt, the header dist 36 + line 9.7 lands the body at 70.1
    expect(phantomPt).toBeCloseTo(24.5, 0)
  })
})
