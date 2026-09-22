/**
 * Element-level paragraph format (bullet/line spacing/paragraph spacing/alignment):
 * setElementParagraphFormat + surgical pPr patch (run bytes untouched, unmodeled children like defRPr kept).
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { generateParagraphXml } from '../src/generate'
import { patchedElementXml, setElementParagraphFormat } from '../src/index'
import { parseMasterTextStyles } from '../src/placeholder'
import type { TextElement } from '../src/types'

const slideWith = (sp: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${sp}</p:spTree></p:cSld></p:sld>`
const spWith = (txBodyInner: string) =>
  '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody>${txBodyInner}</p:txBody></p:sp>`

const parseOne = (txBodyInner: string) => {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: slideWith(spWith(txBodyInner)),
    ctx: {},
  })
  return { slide, el: slide.elements[0] as TextElement }
}

describe('setElementParagraphFormat', () => {
  it('add bullet: buChar + hanging indent injected, run bytes untouched', () => {
    const RUN = '<a:r><a:rPr sz="1800" b="1"/><a:t>Hello</a:t></a:r>'
    const { slide, el } = parseOne(
      `<a:bodyPr/><a:p>${RUN}</a:p><a:p><a:r><a:t>Two</a:t></a:r></a:p>`,
    )
    expect(setElementParagraphFormat(slide, el.id, { bullet: 'char' })).toBe(true)
    const out = patchedElementXml(el)
    expect(out.match(/<a:buChar char="•"\/>/g)!.length).toBe(2)
    expect(out).toContain('marL="285750"')
    expect(out).toContain('indent="-285750"')
    expect(out).toContain(RUN) // original run bytes fully preserved
  })

  it('gallery preset: buFont written before buChar; a plain character drops the symbol font', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    setElementParagraphFormat(slide, el.id, {
      bullet: 'char',
      bulletChar: '§',
      bulletFont: 'Wingdings',
    })
    expect(patchedElementXml(el)).toContain('<a:buFont typeface="Wingdings"/><a:buChar char="§"/>')
    setElementParagraphFormat(slide, el.id, { bullet: 'char', bulletChar: '•' })
    const out = patchedElementXml(el)
    expect(out).not.toContain('buFont')
    expect(out).toContain('<a:buChar char="•"/>')
  })

  it('picture bullet keeps its blip relationship through a size change', () => {
    const { slide, el } = parseOne(
      '<a:bodyPr/><a:p><a:pPr marL="342900" indent="-342900"><a:buBlip><a:blip r:embed="rId6"/></a:buBlip></a:pPr><a:r><a:t>x</a:t></a:r></a:p>',
    )
    setElementParagraphFormat(slide, el.id, { bulletSizePct: 150 })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buSzPct val="150000"/><a:buBlip><a:blip r:embed="rId6"/></a:buBlip>')
    expect(out.match(/<a:buBlip>/g)!.length).toBe(1) // the original element is lifted, not duplicated
  })

  it('picture bullet inherited from the master: a size change writes no glyph, the picture keeps inheriting', () => {
    const master =
      '<?xml version="1.0"?><p:sldMaster xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>' +
      '<p:txStyles><p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900"><a:buBlip><a:blip r:embed="rId9"/></a:buBlip></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>'
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:nvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
          '<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>',
      ),
      ctx: {
        masterTextStyles: parseMasterTextStyles(
          master,
          undefined,
          new Map([['rId9', 'ppt/media/m.png']]),
        ),
      },
    })
    const el = slide.elements[0] as TextElement
    expect(el.text!.paragraphs[0]!.bullet).toEqual({ type: 'blip', mediaRef: 'ppt/media/m.png' })
    setElementParagraphFormat(slide, el.id, { bulletSizePct: 150 })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buSzPct val="150000"/>')
    expect(out).not.toMatch(/buBlip|buChar|rId9/)
  })

  it('custom bullet character written as buChar', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    setElementParagraphFormat(slide, el.id, { bullet: 'char', bulletChar: '◆' })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buChar char="◆"/>')
    expect(el.text!.paragraphs[0]!.bullet).toEqual({ type: 'char', char: '◆' })
  })

  it('bulletHangEmu adjusts hanging indent of existing bullets', () => {
    const { slide, el } = parseOne(
      '<a:bodyPr/><a:p><a:pPr marL="342900" indent="-342900"><a:buChar char="•"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>' +
        '<a:p><a:r><a:t>plain</a:t></a:r></a:p>',
    )
    setElementParagraphFormat(slide, el.id, { bulletHangEmu: 114300 })
    const out = patchedElementXml(el)
    expect(out).toContain('marL="114300"')
    expect(out).toContain('indent="-114300"')
    // The bullet-less paragraph is untouched
    expect(el.text!.paragraphs[1]!.marL).toBeUndefined()
  })

  it('bulletHangEmu together with bullet overrides the default hang', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    setElementParagraphFormat(slide, el.id, {
      bullet: 'char',
      bulletChar: '-',
      bulletHangEmu: 342900,
    })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buChar char="-"/>')
    expect(out).toContain('marL="342900"')
    expect(out).toContain('indent="-342900"')
  })

  it('bullet size/color written as buSzPct/buClr, bullet-less paragraph untouched', () => {
    const { slide, el } = parseOne(
      '<a:bodyPr/><a:p><a:pPr marL="228600" indent="-228600"><a:buChar char="•"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>' +
        '<a:p><a:r><a:t>plain</a:t></a:r></a:p>',
    )
    setElementParagraphFormat(slide, el.id, { bulletSizePct: 80, bulletColor: '#C00000' })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buClr><a:srgbClr val="C00000"/></a:buClr>')
    expect(out).toContain('<a:buSzPct val="80000"/>')
    // Schema order: buClr → buSzPct → buChar
    expect(out).toMatch(/<a:buClr>[\s\S]*?<a:buSzPct[^>]*\/><a:buChar/)
    expect(el.text!.paragraphs[0]!.bullet).toMatchObject({ sizePct: 80, color: '#C00000' })
    expect(el.text!.paragraphs[1]!.bullet).toBeUndefined()
  })

  it('a new bullet color replaces a captured theme buClr node (no stale schemeClr)', () => {
    const theme = { colors: { accent1: '4472C4', dk1: '000000', lt1: 'FFFFFF' } } as any
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(
        spWith(
          '<a:bodyPr/><a:p><a:pPr marL="228600" indent="-228600">' +
            '<a:buClr><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:buClr>' +
            '<a:buChar char="•"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>',
        ),
      ),
      ctx: { theme },
    })
    const el = slide.elements[0] as TextElement
    expect(el.text!.paragraphs[0]!.bullet!.colorNodeXml).toContain('schemeClr')
    setElementParagraphFormat(slide, el.id, { bulletColor: '#0070C0' })
    expect(el.text!.paragraphs[0]!.bullet!.colorNodeXml).toBeUndefined()
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buClr><a:srgbClr val="0070C0"/></a:buClr>')
    expect(out).not.toContain('schemeClr')
  })

  it('changing the bullet glyph keeps existing color/size', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    setElementParagraphFormat(slide, el.id, { bullet: 'char' })
    setElementParagraphFormat(slide, el.id, { bulletSizePct: 125, bulletColor: '#0070C0' })
    setElementParagraphFormat(slide, el.id, { bullet: 'char', bulletChar: '◆' })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buChar char="◆"/>')
    expect(out).toContain('<a:buSzPct val="125000"/>')
    expect(out).toContain('<a:buClr><a:srgbClr val="0070C0"/></a:buClr>')
  })

  it('numbered list writes buAutoNum', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    setElementParagraphFormat(slide, el.id, { bullet: 'number' })
    expect(patchedElementXml(el)).toContain('<a:buAutoNum type="arabicPeriod"/>')
  })

  it('numbering scheme and start number written on buAutoNum; a scheme change alone re-schemes numbered paragraphs only', () => {
    const { slide, el } = parseOne(
      '<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p><a:p><a:r><a:t>y</a:t></a:r></a:p>',
    )
    setElementParagraphFormat(
      slide,
      el.id,
      { bullet: 'number', numType: 'romanUcPeriod', startAt: 4 },
      [0],
    )
    setElementParagraphFormat(slide, el.id, { bullet: 'char' }, [1])
    expect(patchedElementXml(el)).toContain('<a:buAutoNum type="romanUcPeriod" startAt="4"/>')
    setElementParagraphFormat(slide, el.id, { numType: 'alphaLcParenR' })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buAutoNum type="alphaLcParenR" startAt="4"/>')
    expect(out.match(/<a:buChar char="•"\/>/g)!.length).toBe(1)
    // Re-applying plain numbering keeps the scheme and start
    setElementParagraphFormat(slide, el.id, { bullet: 'number' }, [0])
    expect(patchedElementXml(el)).toContain('<a:buAutoNum type="alphaLcParenR" startAt="4"/>')
  })

  it('picture bullet: a landed media rel is written as buBlip with the hanging indent', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    setElementParagraphFormat(slide, el.id, {
      bullet: 'blip',
      bulletBlip: { mediaRef: 'ppt/media/image7.png', blipEmbedId: 'rId9' },
    })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buBlip><a:blip r:embed="rId9"/></a:buBlip>')
    expect(out).toContain('indent="-285750"')
    expect(el.text!.paragraphs[0]!.bullet).toMatchObject({
      type: 'blip',
      mediaRef: 'ppt/media/image7.png',
    })
  })

  it('remove bullet: buNone + indent reset to zero; existing explicit buChar replaced', () => {
    const { slide, el } = parseOne(
      '<a:bodyPr/><a:p><a:pPr marL="342900" indent="-342900"><a:buChar char="v"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>',
    )
    setElementParagraphFormat(slide, el.id, { bullet: 'none' })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buNone/>')
    expect(out).not.toContain('buChar')
    expect(out).toContain('marL="0"')
    expect(out).toContain('indent="0"')
  })

  it('line/paragraph spacing written, unmodeled defRPr inside pPr preserved', () => {
    const { slide, el } = parseOne(
      '<a:bodyPr/><a:p><a:pPr algn="ctr"><a:defRPr sz="2000"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>',
    )
    setElementParagraphFormat(slide, el.id, { lineSpacingPct: 150, spaceAfterPt: 12 })
    const out = patchedElementXml(el)
    expect(out).toContain('<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>')
    expect(out).toContain('<a:spcAft><a:spcPts val="1200"/></a:spcAft>')
    expect(out).toContain('<a:defRPr sz="2000"/>') // unmodeled child elements untouched
    expect(out).toContain('algn="ctr"') // unchanged attributes untouched
  })

  it('justify + replace existing lnSpc (schema order keeps lnSpc first)', () => {
    const { slide, el } = parseOne(
      '<a:bodyPr/><a:p><a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr><a:r><a:t>x</a:t></a:r></a:p>',
    )
    setElementParagraphFormat(slide, el.id, { align: 'justify', lineSpacingPct: 200 })
    const out = patchedElementXml(el)
    expect(out).toContain('algn="just"')
    expect(out).toContain('<a:lnSpc><a:spcPct val="200000"/></a:lnSpc>')
    expect(out).toContain('<a:spcBef><a:spcPts val="600"/></a:spcBef>') // unchanged group preserved
    expect(out).toMatch(/<a:pPr[^>]*><a:lnSpc>[\s\S]*?<a:spcBef>/) // order is schema-valid
    expect(out.match(/<a:lnSpc>/g)!.length).toBe(1)
  })

  it('model-level pPrExplicit flags set (later rebuild path also writes them back)', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    setElementParagraphFormat(slide, el.id, { bullet: 'char', lineSpacingPct: 115 })
    const p = el.text!.paragraphs[0]!
    expect(p.pPrExplicit?.bullet).toBe(true)
    expect(p.pPrExplicit?.lnSpc).toBe(true)
    expect(p.bullet?.type).toBe('char')
    expect(p.lineHeight).toBe(115)
  })
})

describe('paragraph direction rtl', () => {
  it('rtl=true injects rtl="1" on pPr, run bytes untouched', () => {
    const RUN = '<a:r><a:rPr sz="1800"/><a:t>Hi</a:t></a:r>'
    const { slide, el } = parseOne(`<a:bodyPr/><a:p>${RUN}</a:p>`)
    expect(setElementParagraphFormat(slide, el.id, { rtl: true })).toBe(true)
    const out = patchedElementXml(el)
    expect(out).toContain('rtl="1"')
    expect(out).toContain(RUN)
    expect(el.text!.paragraphs[0]!.rtl).toBe(true)
  })

  it('rtl=false rewrites an existing rtl="1" to rtl="0" (explicit LTR base)', () => {
    const { slide, el } = parseOne(
      '<a:bodyPr/><a:p><a:pPr rtl="1" algn="r"/><a:r><a:t>x</a:t></a:r></a:p>',
    )
    setElementParagraphFormat(slide, el.id, { rtl: false })
    const out = patchedElementXml(el)
    expect(out).toContain('rtl="0"')
    expect(out).not.toContain('rtl="1"')
    expect(out).toContain('algn="r"') // untouched sibling attribute
    expect(el.text!.paragraphs[0]!.rtl).toBe(false)
  })
})

describe('multi-level indentation indentDelta', () => {
  it('increasing level writes lvl, own bullet hanging indent grows with the level', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    setElementParagraphFormat(slide, el.id, { bullet: 'char' })
    setElementParagraphFormat(slide, el.id, { indentDelta: 1 })
    const out = patchedElementXml(el)
    expect(out).toContain('lvl="1"')
    expect(out).toContain('marL="571500"') // 285750 * (1+1)
    setElementParagraphFormat(slide, el.id, { indentDelta: -1 })
    const out2 = patchedElementXml(el)
    expect(out2).not.toContain('lvl=')
    expect(out2).toContain('marL="285750"')
  })

  it('level clamped between 0 and 8', () => {
    const { slide, el } = parseOne('<a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p>')
    expect(setElementParagraphFormat(slide, el.id, { indentDelta: -1 })).toBe(false)
    for (let i = 0; i < 12; i++) setElementParagraphFormat(slide, el.id, { indentDelta: 1 })
    expect(el.text!.paragraphs[0]!.level).toBe(8)
  })
})

describe('per-paragraph paraIndices (editing-mode selection)', () => {
  const TWO_PARAS =
    '<a:bodyPr/><a:p><a:r><a:t>Heading</a:t></a:r></a:p>' +
    '<a:p><a:pPr marL="228600" indent="-228600"><a:buChar char="•"/></a:pPr><a:r><a:t>Body</a:t></a:r></a:p>'

  it('bullet on paragraph 1 only: paragraph 0 bytes untouched', () => {
    const { slide, el } = parseOne(TWO_PARAS)
    expect(setElementParagraphFormat(slide, el.id, { bullet: 'char', bulletChar: '○' }, [1])).toBe(
      true,
    )
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buChar char="○"/>')
    expect(out).toContain('<a:p><a:r><a:t>Heading</a:t></a:r></a:p>') // p0 verbatim, no pPr injected
    expect(el.text!.paragraphs[0]!.bullet).toBeUndefined()
    expect(el.dirtyPPr?.paraIndices).toEqual([1])
  })

  it('bullet off on paragraph 1 only (the issue scenario: heading stays clean)', () => {
    const { slide, el } = parseOne(TWO_PARAS)
    setElementParagraphFormat(slide, el.id, { bullet: 'none' }, [1])
    const out = patchedElementXml(el)
    expect(out).toContain('<a:buNone/>')
    expect(out).toContain('<a:p><a:r><a:t>Heading</a:t></a:r></a:p>')
  })

  it('restricted ops accumulate indices; an element-wide op widens to all paragraphs', () => {
    const { slide, el } = parseOne(TWO_PARAS)
    setElementParagraphFormat(slide, el.id, { lineSpacingPct: 150 }, [0])
    setElementParagraphFormat(slide, el.id, { spaceAfterPt: 6 }, [1])
    expect(el.dirtyPPr?.paraIndices).toEqual([0, 1])
    setElementParagraphFormat(slide, el.id, { lineSpacingPct: 115 })
    expect(el.dirtyPPr?.paraIndices).toBeUndefined()
  })

  it('out-of-range indices are a no-op', () => {
    const { slide, el } = parseOne(TWO_PARAS)
    expect(setElementParagraphFormat(slide, el.id, { bullet: 'char' }, [9])).toBe(false)
    expect(el.dirtyPPr).toBeUndefined()
  })
})

describe('empty paragraph endParaRPr', () => {
  it('empty paragraph takes its line style from endParaRPr (marker run)', () => {
    const { el } = parseOne(
      '<a:bodyPr/><a:p><a:r><a:rPr sz="2000"/><a:t>X</a:t></a:r></a:p>' +
        '<a:p><a:endParaRPr sz="8000" b="1"/></a:p>',
    )
    const p = el.text!.paragraphs[1]!
    expect(p.runs.length).toBe(1)
    expect(p.runs[0]!.text).toBe('')
    expect(p.runs[0]!.fontSize).toBe(80)
    expect(p.runs[0]!.bold).toBe(true)
  })

  it("endParaRPr overrides an empty run's own rPr (probe-measured PowerPoint rule)", () => {
    const { el } = parseOne(
      '<a:bodyPr/><a:p><a:r><a:rPr sz="1400"/><a:t></a:t></a:r>' +
        '<a:endParaRPr sz="8000"/></a:p>',
    )
    const p = el.text!.paragraphs[0]!
    expect(p.runs.length).toBe(1)
    expect(p.runs[0]!.fontSize).toBe(80)
  })

  it('a paragraph with real text keeps its runs (endParaRPr ignored)', () => {
    const { el } = parseOne(
      '<a:bodyPr/><a:p><a:r><a:rPr sz="2000"/><a:t>Hello</a:t></a:r>' +
        '<a:endParaRPr sz="8000"/></a:p>',
    )
    const p = el.text!.paragraphs[0]!
    expect(p.runs[0]!.text).toBe('Hello')
    expect(p.runs[0]!.fontSize).toBe(20)
  })
})

describe('East Asian wrap flags', () => {
  it('eaLnBrk/latinLnBrk/hangingPunct parse from pPr and are written back by the generator', () => {
    const { el } = parseOne(
      '<a:bodyPr/><a:p><a:pPr eaLnBrk="0" latinLnBrk="1" hangingPunct="0"/><a:r><a:t>Hi</a:t></a:r></a:p>',
    )
    const p = el.text!.paragraphs[0]!
    expect(p.eaLnBrk).toBe(false)
    expect(p.latinLnBrk).toBe(true)
    expect(p.hangingPunct).toBe(false)
    const out = generateParagraphXml(p)
    expect(out).toContain('eaLnBrk="0"')
    expect(out).toContain('latinLnBrk="1"')
    expect(out).toContain('hangingPunct="0"')
    // defaults stay implicit
    expect(generateParagraphXml({ runs: [{ text: 'a' }] })).not.toMatch(
      /eaLnBrk|latinLnBrk|hangingPunct/,
    )
  })
})
