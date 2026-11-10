/**
 * Patches that used to leave the slide XML schema-invalid (PowerPoint "needs repair"):
 *  - a run color on a rPr that already carries noFill/gradFill/solidFill(non-srgb) → two fills
 *  - a run color on WordArt (a:ln first) recolored the outline instead of the text
 *  - sz below ST_TextFontSize's floor
 *  - a:ln appended after effectLst when spPr has no fill/geometry/xfrm
 *  - edits on an mc:AlternateContent-anchored shape spliced across the Choice/Fallback boundary
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { XMLValidator } from 'fast-xml-parser'
import { validatePptx, xmllintAvailable } from '../../../tools/ooxml-validate/validate-pptx.mjs'
import {
  openPptx,
  savePptx,
  addElement,
  createBlankPptx,
  setElementFont,
  setElementFill,
  setElementParagraphFormat,
  strokePatchToModel,
} from '../src/index'

const GEOM =
  '<a:xfrm><a:off x="100" y="100"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
const sp = (name: string, body: string, spPr = GEOM) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="90" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr>${spPr}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`
const run = (rPr: string, t = 'hello') =>
  `<a:r><a:rPr lang="en-US"${rPr ? `>${rPr}</a:rPr>` : '/>'}<a:t>${t}</a:t></a:r>`
const alt = (choice: string, fallback: string) =>
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  `<mc:Choice Requires="a14" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">${choice}</mc:Choice>` +
  `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`

async function deckWith(fragment: string) {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  addElement(slide, {
    kind: 'rect',
    offset: { x: 0, y: 0, cx: 100, cy: 100 },
    fillColor: '#111111',
  })
  const zip = await JSZip.loadAsync(await savePptx(opened))
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const at = xml.indexOf('</p:sp>') + '</p:sp>'.length
  zip.file('ppt/slides/slide1.xml', xml.slice(0, at) + fragment + xml.slice(at))
  const reopened = await openPptx(await zip.generateAsync({ type: 'nodebuffer' }))
  const s = reopened.deck.slides[0]!
  return { opened: reopened, slide: s, el: s.elements[s.elements.length - 1]! }
}

const schemaGate = xmllintAvailable() || !!process.env.CI

async function savedSlideXml(opened: Awaited<ReturnType<typeof openPptx>>) {
  const saved = await savePptx(opened)
  const zip = await JSZip.loadAsync(saved)
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  expect(XMLValidator.validate(xml)).toBe(true)
  if (schemaGate) expect(await validatePptx(saved)).toEqual([])
  return xml
}

const probeRPr = (xml: string) =>
  /name="probe"[\s\S]*?<a:rPr\b[^>]*>([\s\S]*?)<\/a:rPr>/.exec(xml)![1]!
const fillCount = (rPr: string) =>
  (rPr.match(/<a:(?:noFill|solidFill|gradFill|pattFill|blipFill|grpFill)\b/g) ?? []).length

describe('run color keeps one fill child in a:rPr', () => {
  it('replaces a gradFill text fill instead of adding a second fill', async () => {
    const { opened, slide, el } = await deckWith(
      sp(
        'probe',
        `<a:p>${run('<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill><a:latin typeface="Arial"/>')}</a:p>`,
      ),
    )
    expect(setElementFont(slide, el.id, { color: '#FFFFFF' })).toBe(true)
    const rPr = probeRPr(await savedSlideXml(opened))
    expect(fillCount(rPr)).toBe(1)
    expect(rPr).toContain('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin')
  })

  it('replaces noFill and a solidFill/prstClr', async () => {
    for (const existing of ['<a:noFill/>', '<a:solidFill><a:prstClr val="black"/></a:solidFill>']) {
      const { opened, slide, el } = await deckWith(sp('probe', `<a:p>${run(existing)}</a:p>`))
      expect(setElementFont(slide, el.id, { color: '#FFFFFF' })).toBe(true)
      const rPr = probeRPr(await savedSlideXml(opened))
      expect(fillCount(rPr)).toBe(1)
      expect(rPr).toContain('<a:srgbClr val="FFFFFF"/>')
    }
  })

  it('recolors the text fill of WordArt, not its a:ln outline, and injects after a:ln', async () => {
    const outlined = await deckWith(
      sp(
        'probe',
        `<a:p>${run('<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>')}</a:p>`,
      ),
    )
    expect(setElementFont(outlined.slide, outlined.el.id, { color: '#FFFFFF' })).toBe(true)
    const rPr = probeRPr(await savedSlideXml(outlined.opened))
    expect(rPr).toBe(
      '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
    )

    const lnOnly = await deckWith(
      sp(
        'probe',
        `<a:p>${run('<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:latin typeface="Arial"/>')}</a:p>`,
      ),
    )
    expect(setElementFont(lnOnly.slide, lnOnly.el.id, { color: '#FFFFFF' })).toBe(true)
    expect(probeRPr(await savedSlideXml(lnOnly.opened))).toBe(
      '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/>',
    )
  })

  it('keeps the alpha modifier of an existing srgbClr', async () => {
    const { opened, slide, el } = await deckWith(
      sp(
        'probe',
        `<a:p>${run('<a:solidFill><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>')}</a:p>`,
      ),
    )
    expect(setElementFont(slide, el.id, { color: '#FFFFFF' })).toBe(true)
    expect(probeRPr(await savedSlideXml(opened))).toBe(
      '<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
    )
  })
})

describe('font size stays inside ST_TextFontSize', () => {
  it('clamps sz to 100..400000', async () => {
    const small = await deckWith(sp('probe', `<a:p>${run('')}</a:p>`))
    expect(setElementFont(small.slide, small.el.id, { fontSizePt: 0.5 })).toBe(true)
    expect(await savedSlideXml(small.opened)).toContain('name="probe"')
    expect(
      /name="probe"[\s\S]*?<a:rPr sz="(\d+)"/.exec(await savedSlideXml(small.opened))![1],
    ).toBe('100')

    const huge = await deckWith(sp('probe', `<a:p>${run('')}</a:p>`))
    expect(setElementFont(huge.slide, huge.el.id, { fontSizePt: 99999 })).toBe(true)
    expect(/name="probe"[\s\S]*?<a:rPr sz="(\d+)"/.exec(await savedSlideXml(huge.opened))![1]).toBe(
      '400000',
    )
  })
})

describe('stroke insertion order', () => {
  it('puts a:ln before effectLst when spPr has no fill, geometry or xfrm', async () => {
    const { opened, el } = await deckWith(
      sp(
        'probe',
        `<a:p>${run('')}</a:p>`,
        '<a:effectLst><a:outerShdw dist="1"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>',
      ),
    )
    ;(el as any).stroke = strokePatchToModel({ color: '#FF0000', widthEmu: 12700 })
    el.dirtyStroke = true
    const xml = await savedSlideXml(opened)
    const spPr = /name="probe"[\s\S]*?<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(xml)![1]!
    expect(spPr.indexOf('<a:ln')).toBeGreaterThanOrEqual(0)
    expect(spPr.indexOf('<a:ln')).toBeLessThan(spPr.indexOf('<a:effectLst'))
  })
})

describe('mc:AlternateContent-anchored shapes', () => {
  const choice = sp('choice', `<a:p>${run('')}${run('', 'world')}</a:p>`)
  const fallback = sp('probe', `<a:p>${run('')}</a:p>`)

  it('color edits land inside the Fallback branch only; the block stays well-formed', async () => {
    const { opened, slide, el } = await deckWith(alt(choice, fallback))
    expect(el.anchor.originalXml.startsWith('<mc:AlternateContent')).toBe(true)
    expect(setElementFont(slide, el.id, { color: '#FFFFFF' })).toBe(true)
    const xml = await savedSlideXml(opened)
    const block = xml.slice(
      xml.indexOf('<mc:AlternateContent'),
      xml.indexOf('</mc:AlternateContent>'),
    )
    const [choiceXml, fallbackXml] = block.split('<mc:Fallback>')
    expect(choiceXml).toContain('<a:t>world</a:t>')
    expect(choiceXml).not.toContain('FFFFFF')
    expect(fallbackXml).toContain('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>')
  })

  it('paragraph format on a math Choice (no runs) does not splice across branches', async () => {
    const math = sp('choice', '<a:p><a14:m/><a:endParaRPr lang="en-US"/></a:p><a:p/>')
    const { opened, slide, el } = await deckWith(alt(math, fallback))
    expect(setElementParagraphFormat(slide, el.id, { align: 'center' })).toBe(true)
    const xml = await savedSlideXml(opened)
    expect(xml).toContain('<a14:m/>')
    expect((xml.match(/algn="ctr"/g) ?? []).length).toBe(1)
  })

  it('shape fill is mirrored onto the same-tag Choice so PowerPoint shows it too', async () => {
    const { opened, slide, el } = await deckWith(alt(choice, fallback))
    expect(setElementFill(opened, slide, el.id, '#FF0000')).toBe(true)
    const xml = await savedSlideXml(opened)
    expect(
      (xml.match(/<a:solidFill><a:srgbClr val="FF0000"\/><\/a:solidFill>/g) ?? []).length,
    ).toBe(2)
  })
})
