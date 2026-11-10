/**
 * Second sweep of byte-patch paths that could leave a saved deck schema-invalid or malformed
 * (every case validated end-to-end with tools/ooxml-validate):
 *  - table style regex swallowing cells between a self-closing and a paired a:tcPr
 *  - effectLst / hlinkClick anchored on an extLst that belongs to a:ln
 *  - `$&` in user strings re-expanded by String.replace (rels, comments, notes, sections)
 *  - unrounded / out-of-range xfrm, line width, paragraph metrics, advTm
 *  - non-hex colors, unknown preset names, slide ids at the ST_SlideId ceiling
 *  - group child scanning descending into mc:AlternateContent / self-closing grpSpPr
 *  - the gate itself: malformed raw parts and .rels are seen
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { validatePptx, xmllintAvailable } from '../../../tools/ooxml-validate/validate-pptx.mjs'
import {
  openPptx,
  savePptx,
  addElement,
  addTable,
  addSection,
  addSlideComment,
  createBlankPptx,
  editTableStyle,
  groupChildSlices,
  setElementEffects,
  setElementFill,
  setElementFont,
  setElementLink,
  setElementParagraphFormat,
  setShapePresetGeometry,
  setSlideAdvanceTime,
  setSlideNotes,
  strokePatchToModel,
  type TextElement,
} from '../src/index'

const available = xmllintAvailable()

async function blankWithShape(spPrExtra = '', rPrInner = '') {
  const opened = await openPptx(await createBlankPptx())
  addElement(opened.deck.slides[0]!, {
    kind: 'rect',
    offset: { x: 100, y: 100, cx: 1000000, cy: 500000 },
    paragraphs: [{ runs: [{ text: 'probe' }] }],
  })
  const zip = await JSZip.loadAsync(await savePptx(opened))
  let xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  if (spPrExtra) xml = xml.replace('</p:spPr>', `${spPrExtra}</p:spPr>`)
  if (rPrInner) xml = xml.replace(/<a:rPr\b([^>]*?)\/>/, `<a:rPr$1>${rPrInner}</a:rPr>`)
  zip.file('ppt/slides/slide1.xml', xml)
  const reopened = await openPptx(await zip.generateAsync({ type: 'uint8array' }))
  const slide = reopened.deck.slides[0]!
  return { opened: reopened, slide, el: slide.elements[slide.elements.length - 1]! }
}

async function saveClean(opened: Awaited<ReturnType<typeof openPptx>>) {
  const saved = await savePptx(opened)
  if (available || process.env.CI) expect(await validatePptx(saved)).toEqual([])
  const zip = await JSZip.loadAsync(saved)
  return { zip, slide1: await zip.file('ppt/slides/slide1.xml')!.async('string') }
}

const LN_WITH_EXT =
  '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
  '<a:extLst><a:ext uri="{C807C97D-BFC1-408E-A445-0C87EB9F89A2}"><a14:hiddenLine xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a14:hiddenLine></a:ext></a:extLst></a:ln>'

describe.skipIf(!available && !process.env.CI)('schema-safe patches, wave 2', () => {
  it('table style edits do not swallow the cells between a self-closing and a paired tcPr', async () => {
    const opened = await openPptx(await createBlankPptx())
    addTable(opened, 0, { rows: 1, cols: 3, offset: { x: 0, y: 0, cx: 3000000, cy: 400000 } })
    const zip = await JSZip.loadAsync(await savePptx(opened))
    let xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    // cell 0: self-closing tcPr + colored run; cell 1: paired tcPr with a fill
    const cells = [...xml.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)].map((m) => m[0])
    expect(cells.length).toBe(3)
    const c0 = cells[0]!.replace(
      '<a:p/>',
      '<a:p><a:r><a:rPr lang="en-US"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>hi</a:t></a:r></a:p>',
    )
    const c1 = cells[1]!.replace(
      /<a:tcPr\b[^>]*\/>|<a:tcPr\b[^>]*>[\s\S]*?<\/a:tcPr>/,
      '<a:tcPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:tcPr>',
    )
    xml = xml.replace(cells[0]!, c0).replace(cells[1]!, c1)
    zip.file('ppt/slides/slide1.xml', xml)
    const reopened = await openPptx(await zip.generateAsync({ type: 'uint8array' }))
    const slide = reopened.deck.slides[0]!
    const table = slide.elements.find((e) => e.type === 'table')!
    expect(
      editTableStyle(slide, table.id, {
        borderColor: '#0000FF',
        borderWidthEmu: 12700,
        borderPreset: 'all',
      }),
    ).toBe(true)
    const { slide1 } = await saveClean(reopened)
    expect(slide1).not.toMatch(/<a:rPr\b[^>]*>(?:(?!<\/a:rPr>)[\s\S])*<a:lnL/)
    expect((slide1.match(/<a:lnL\b/g) ?? []).length).toBe(3)
    expect(slide1).toContain('<a:t>hi</a:t>')
  })

  it('effectLst lands in spPr, not inside an a:ln that carries its own extLst', async () => {
    const { opened, slide, el } = await blankWithShape(LN_WITH_EXT)
    expect(
      setElementEffects(slide, el.id, {
        shadow: { color: '#00000080', blurRad: 50800, dist: 38100, dirDeg: 45 },
      }),
    ).toBe(true)
    const { slide1 } = await saveClean(opened)
    expect(slide1).toMatch(/<\/a:ln><a:effectLst>/)
  })

  it('hlinkClick lands among the rPr children, not inside an a:ln extLst', async () => {
    const { opened, el } = await blankWithShape('', LN_WITH_EXT)
    const run = (el as TextElement).text!.paragraphs[0]!.runs[0]!
    run.hyperlinkRId = 'rId9'
    el.dirty = true
    const { slide1 } = await saveClean(opened)
    expect(slide1).toMatch(/<\/a:ln><a:hlinkClick\b[^>]*r:id="rId9"/)
  })

  it('`$&` in user strings survives String.replace (rels, comments, notes, sections)', async () => {
    const { opened, slide, el } = await blankWithShape()
    expect(
      setElementLink(opened, 0, el.id, { kind: 'url', url: 'https://ex.com/?a=$&b=2' }),
    ).toBeTruthy()
    expect(addSlideComment(opened, 0, { author: 'A $& B', text: 'x $` y' })).toBeTruthy()
    // second comment by the same author takes the author-reuse path
    expect(addSlideComment(opened, 0, { author: 'A $& B', text: 'again' })).toBeTruthy()
    expect(setSlideNotes(opened, 0, "note $' tail")).toBe(true)
    expect(addSection(opened, 0, 'Sec $&')).toBeTruthy()
    void slide
    const { zip } = await saveClean(opened)
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
    expect(rels).toContain('Target="https://ex.com/?a=$&amp;b=2"')
    expect(await zip.file('ppt/presentation.xml')!.async('string')).toContain('name="Sec $&amp;"')
    expect(await zip.file('ppt/commentAuthors.xml')!.async('string')).toMatch(
      /name="A \$&amp; B"[^>]*lastIdx="2"/,
    )
  })

  it('transform, line width and advTm are written as clamped integers', async () => {
    const { opened, slide, el } = await blankWithShape()
    el.transform.offset = { x: 100.5, y: -3.2, cx: 999.7, cy: -5 }
    el.transform.rot = 1e12
    el.dirtyTransform = true
    ;(el as TextElement).stroke = strokePatchToModel({ color: '#FF0000', widthEmu: -5000 })
    el.dirtyStroke = true
    setSlideAdvanceTime(slide, 1500.7)
    const { slide1 } = await saveClean(opened)
    expect(slide1).toContain('<a:off x="101" y="-3"/><a:ext cx="1000" cy="0"/>')
    expect(slide1).toContain('<a:ln w="0">')
    expect(slide1).toContain('advTm="1501"')
  })

  it('paragraph metrics are clamped to their schema ranges', async () => {
    const { opened, slide, el } = await blankWithShape()
    expect(
      setElementParagraphFormat(slide, el.id, {
        bullet: 'char',
        bulletSizePct: 10,
        lineSpacingPct: 20000,
        spaceBeforePt: 2000,
        bulletHangEmu: 1e9,
      }),
    ).toBe(true)
    const { slide1 } = await saveClean(opened)
    expect(slide1).toContain('<a:buSzPct val="25000"/>')
    expect(slide1).toContain('<a:spcPct val="13200000"/>')
    expect(slide1).toContain('<a:spcPts val="158400"/>')
    expect(slide1).toContain('marL="51206400"')
    expect(slide1).toContain('indent="-51206400"')
  })

  it('non-hex colors and unknown preset names never reach the XML', async () => {
    const { opened, slide, el } = await blankWithShape()
    expect(setShapePresetGeometry(slide, el.id, 'bogusShape')).toBe(false)
    expect(setShapePresetGeometry(slide, el.id, 'roundRect')).toBe(true)
    expect(setElementFill(opened, slide, el.id, 'red')).toBe(true)
    expect(setElementFont(slide, el.id, { color: '#f00' })).toBe(true)
    const { slide1 } = await saveClean(opened)
    expect(slide1).toContain('<a:prstGeom prst="roundRect">')
    expect(slide1).toMatch(/<p:spPr>[\s\S]*?<a:solidFill><a:srgbClr val="000000"\/>/)
    expect(slide1).toMatch(/<a:rPr\b[^>]*><a:solidFill><a:srgbClr val="FF0000"\/>/)
  })

  it('group child scanning stops at a self-closing grpSpPr and does not enter mc:AlternateContent', () => {
    const sp = (id: number) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`
    const grp =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      sp(22) +
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="a14">' +
      sp(21) +
      '</mc:Choice><mc:Fallback>' +
      sp(21) +
      '</mc:Fallback></mc:AlternateContent>' +
      sp(23) +
      '</p:grpSp>'
    expect(groupChildSlices(grp).map((s) => s.nvId)).toEqual(['22', '23'])
  })

  it('the gate sees malformed raw parts and .rels', async () => {
    const { opened } = await blankWithShape()
    const zip = await JSZip.loadAsync(await savePptx(opened))
    zip.file(
      'ppt/slides/slide1.xml',
      (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
        '</a:t>',
        '</a:t></a:oops>',
      ),
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      (await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).replace(
        '</Relationships>',
        '<Relationship Id="rId9" Target="x<y"/></Relationships>',
      ),
    )
    const problems = await validatePptx(await zip.generateAsync({ type: 'uint8array' }))
    expect(problems.some((p) => p.part === 'ppt/slides/slide1.xml')).toBe(true)
    expect(problems.some((p) => p.part === 'ppt/slides/_rels/slide1.xml.rels')).toBe(true)
  })
})
