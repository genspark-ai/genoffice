import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import JSZip from 'jszip'
import {
  openPptx,
  savePptx,
  addElement,
  addPicture,
  createBlankPptx,
  editPictureSrcRect,
  patchElementStroke,
  replacePictureBytes,
  setSlideBackground,
} from '../src/index'
import type { PictureElement, TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 1828800, cy: 914400 }

// 1x1 red PNG
const PNG_1PX = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

describe('addPicture', () => {
  it('media/rels/content-types all land and survive save → reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addPicture(opened, slide, { bytes: PNG_1PX, ext: 'png', offset: { ...OFF } })
    expect(el).not.toBeNull()
    expect(el!.mediaRef).toBe('ppt/media/image1.png')

    const out = await savePptx(opened)
    const zip = await JSZip.loadAsync(out)
    expect(zip.file('ppt/media/image1.png')).not.toBeNull()
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('Extension="png"')

    const reopened = await openPptx(out)
    const pic = reopened.deck.slides[0]!.elements.find(
      (e) => e.type === 'picture',
    ) as PictureElement
    expect(pic).toBeDefined()
    expect(pic.mediaRef).toBe('ppt/media/image1.png')
    expect(pic.transform.offset).toEqual(OFF)
  })

  it('second picture gets image2 and a distinct rId', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addPicture(opened, slide, { bytes: PNG_1PX, ext: 'png', offset: { ...OFF } })
    const b = addPicture(opened, slide, { bytes: PNG_1PX, ext: 'png', offset: { ...OFF } })
    expect(b!.mediaRef).toBe('ppt/media/image2.png')
    const reopened = await openPptx(await savePptx(opened))
    const pics = reopened.deck.slides[0]!.elements.filter((e) => e.type === 'picture')
    expect(pics.length).toBe(2)
    expect(new Set(pics.map((p) => (p as PictureElement).mediaRef)).size).toBe(2)
  })
})

describe('replacePictureBytes', () => {
  // 1x1 transparent GIF (distinct format exercises the Content_Types Default too)
  const GIF_1PX = Uint8Array.from(
    Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64'),
  )

  async function openWithPicture() {
    const opened0 = await openPptx(await createBlankPptx())
    addPicture(opened0, opened0.deck.slides[0]!, { bytes: PNG_1PX, ext: 'png', offset: { ...OFF } })
    // reopen so the picture is a parsed element with on-disk originalXml
    const opened = await openPptx(await savePptx(opened0))
    const slide = opened.deck.slides[0]!
    const pic = slide.elements.find((e) => e.type === 'picture') as PictureElement
    return { opened, slide, pic }
  }

  it('swaps media in place, keeps frame + crop when asked, survives save → reopen', async () => {
    const { opened, slide, pic } = await openWithPicture()
    expect(editPictureSrcRect(slide, pic.id, { l: 0.1, t: 0.1, r: 0.1, b: 0.1 })).toBe(true)
    expect(replacePictureBytes(opened, slide, pic.id, GIF_1PX, 'gif', { keepSrcRect: true })).toBe(
      true,
    )
    expect(pic.mediaRef).toBe('ppt/media/image2.gif')

    const out = await savePptx(opened)
    const zip = await JSZip.loadAsync(out)
    expect(new Uint8Array(await zip.file('ppt/media/image2.gif')!.async('uint8array'))).toEqual(
      GIF_1PX,
    )
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('Extension="gif"')

    const reopened = await openPptx(out)
    const pic2 = reopened.deck.slides[0]!.elements.find(
      (e) => e.type === 'picture',
    ) as PictureElement
    expect(pic2.mediaRef).toBe('ppt/media/image2.gif')
    expect(pic2.srcRect).toEqual({ l: 0.1, t: 0.1, r: 0.1, b: 0.1 })
    expect(pic2.transform.offset).toEqual(OFF)
  })

  it('clears a stale crop window by default', async () => {
    const { opened, slide, pic } = await openWithPicture()
    editPictureSrcRect(slide, pic.id, { l: 0.2, t: 0, r: 0.2, b: 0 })
    expect(replacePictureBytes(opened, slide, pic.id, GIF_1PX, 'gif')).toBe(true)
    expect(pic.srcRect).toBeUndefined()

    const reopened = await openPptx(await savePptx(opened))
    const pic2 = reopened.deck.slides[0]!.elements.find(
      (e) => e.type === 'picture',
    ) as PictureElement
    expect(pic2.srcRect).toBeUndefined()
  })

  it('keeps stroke and z-order: the element is mutated, not re-added', async () => {
    const { opened, slide, pic } = await openWithPicture()
    pic.stroke = { fill: { type: 'solid', color: '#C43E1C' }, width: 12700 }
    pic.dirtyStroke = true
    addElement(slide, { kind: 'rect', offset: { ...OFF }, fillColor: '#112233' })
    const picIndex = slide.elements.indexOf(pic)
    expect(replacePictureBytes(opened, slide, pic.id, GIF_1PX, 'gif')).toBe(true)
    expect(slide.elements.indexOf(pic)).toBe(picIndex)

    const reopened = await openPptx(await savePptx(opened))
    const els = reopened.deck.slides[0]!.elements
    expect(els.findIndex((e) => e.type === 'picture')).toBe(picIndex)
    const pic2 = els[picIndex] as PictureElement
    expect(pic2.stroke?.fill).toEqual({ type: 'solid', color: '#C43E1C' })
  })

  it('drops a stale svgBlip extension and companion r:link', async () => {
    const { opened, slide, pic } = await openWithPicture()
    // Simulate an SVG picture inserted by Office 2016+: the blip carries a raster
    // fallback (r:embed), an external link (r:link), and the svgBlip extension
    const anchor = (pic as unknown as { anchor: { originalXml: string } }).anchor
    anchor.originalXml = anchor.originalXml.replace(
      /<a:blip r:embed="(rId\d+)"\s*\/>/,
      '<a:blip r:embed="$1" r:link="rId99">' +
        '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
        '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="$1"/>' +
        '</a:ext></a:extLst></a:blip>',
    )
    expect(anchor.originalXml).toContain('svgBlip')
    expect(replacePictureBytes(opened, slide, pic.id, GIF_1PX, 'gif')).toBe(true)
    expect(anchor.originalXml).not.toContain('svgBlip')
    expect(anchor.originalXml).not.toContain('r:link')
    expect(anchor.originalXml).toContain(`r:embed="rId`)
    expect(anchor.originalXml).not.toContain('<a:extLst></a:extLst>')
  })

  it('unknown extension is rejected without touching the element', async () => {
    const { opened, slide, pic } = await openWithPicture()
    expect(replacePictureBytes(opened, slide, pic.id, GIF_1PX, 'exe')).toBe(false)
    expect(pic.mediaRef).toBe('ppt/media/image1.png')
  })
})

describe('patchElementStroke', () => {
  it('stroke color+width round-trips through save → reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'rect', offset: { ...OFF }, fillColor: '#112233' })
    el.stroke = { fill: { type: 'solid', color: '#C43E1C' }, width: 3 * 12700 }
    el.dirtyStroke = true

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    expect(el2.stroke?.fill).toEqual({ type: 'solid', color: '#C43E1C' })
    expect(el2.stroke?.width).toBe(3 * 12700)
  })

  it('picture border round-trips through save → reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const pic = addPicture(opened, slide, { bytes: PNG_1PX, ext: 'png', offset: { ...OFF } })!
    pic.stroke = { fill: { type: 'solid', color: '#C43E1C' }, width: 2 * 12700, dash: 'sysDot' }
    pic.dirtyStroke = true

    const reopened = await openPptx(await savePptx(opened))
    const pic2 = reopened.deck.slides[0]!.elements.find(
      (e) => e.type === 'picture',
    ) as PictureElement
    expect(pic2.stroke?.fill).toEqual({ type: 'solid', color: '#C43E1C' })
    expect(pic2.stroke?.width).toBe(2 * 12700)
    expect(pic2.stroke?.dash).toBe('sysDot')
  })

  it('null stroke writes explicit noFill, keeping other ln bytes', () => {
    const xml =
      '<p:sp><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr></p:sp>'
    const out = patchElementStroke(xml, null)
    expect(out).toContain('<a:ln w="12700"><a:noFill/></a:ln>')
  })

  it('keeps dash/arrows/cap/cmpd when only color+width change', () => {
    const xml =
      '<p:sp><p:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom>' +
      '<a:ln w="9525" cap="rnd" cmpd="dbl" algn="ctr">' +
      '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
      '<a:prstDash val="dash"/><a:round/>' +
      '<a:headEnd type="oval" w="med" len="med"/><a:tailEnd type="arrow" w="lg" len="lg"/>' +
      '</a:ln></p:spPr></p:sp>'
    const out = patchElementStroke(xml, { color: '#C43E1C', widthEmu: 25400 })
    expect(out).toContain('<a:ln w="25400" cap="rnd" cmpd="dbl" algn="ctr">')
    expect(out).toContain('<a:solidFill><a:srgbClr val="C43E1C"/></a:solidFill>')
    expect(out).toContain('<a:prstDash val="dash"/><a:round/>')
    expect(out).toContain('<a:headEnd type="oval" w="med" len="med"/>')
    expect(out).toContain('<a:tailEnd type="arrow" w="lg" len="lg"/>')
    expect(out).not.toContain('val="000000"')
  })

  it('patches a self-closing <a:ln/> and a ln without w attribute', () => {
    const selfClosed = '<p:sp><p:spPr><a:ln cap="sq"/></p:spPr></p:sp>'
    const out1 = patchElementStroke(selfClosed, { color: '#112233', widthEmu: 12700 })
    expect(out1).toContain(
      '<a:ln w="12700" cap="sq"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln>',
    )

    const noW = '<p:sp><p:spPr><a:ln><a:prstDash val="sysDot"/></a:ln></p:spPr></p:sp>'
    const out2 = patchElementStroke(noW, { color: '#445566', widthEmu: 19050 })
    expect(out2).toContain('<a:ln w="19050">')
    // fill is inserted before prstDash (CT_LineProperties sequence)
    expect(out2).toContain(
      '<a:solidFill><a:srgbClr val="445566"/></a:solidFill><a:prstDash val="sysDot"/>',
    )
  })

  it('gradient line fill is replaced, dash preserved', () => {
    const xml =
      '<p:sp><p:spPr><a:ln w="9525">' +
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill>' +
      '<a:prstDash val="lgDash"/></a:ln></p:spPr></p:sp>'
    const out = patchElementStroke(xml, { color: '#00FF00', widthEmu: 9525 })
    expect(out).toContain(
      '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="lgDash"/>',
    )
    expect(out).not.toContain('gradFill')
  })

  it('dash preset writes <a:prstDash> after the fill, replacing an existing one', () => {
    const xml =
      '<p:sp><p:spPr><a:ln w="9525">' +
      '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
      '<a:prstDash val="lgDash"/><a:headEnd type="oval"/></a:ln></p:spPr></p:sp>'
    const out = patchElementStroke(xml, { color: '#C00000', widthEmu: 19050, dash: 'sysDot' })
    expect(out).toContain(
      '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="sysDot"/><a:headEnd type="oval"/>',
    )
    expect(out).not.toContain('lgDash')
  })

  it("dash 'solid' removes the prstDash node; missing ln gets one with a dash", () => {
    const xml =
      '<p:sp><p:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr></p:sp>'
    const out = patchElementStroke(xml, { color: '#000000', widthEmu: 9525, dash: 'solid' })
    expect(out).not.toContain('prstDash')

    const bare = '<p:sp><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'
    const out2 = patchElementStroke(bare, { color: '#C00000', widthEmu: 19050, dash: 'sysDot' })
    expect(out2).toContain(
      '<a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="sysDot"/></a:ln>',
    )
  })

  it('stroke dash round-trips through save → reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'rect', offset: { ...OFF }, fillColor: '#112233' })
    el.stroke = { fill: { type: 'solid', color: '#C43E1C' }, width: 19050, dash: 'sysDot' }
    el.dirtyStroke = true

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    expect(el2.stroke?.dash).toBe('sysDot')
  })
})

describe('setSlideBackground', () => {
  it('injects <p:bg> and round-trips through save → reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideBackground(opened.deck.slides[0]!, '#1A2B3C')
    expect(opened.deck.slides[0]!.background).toEqual({ type: 'solid', color: '#1A2B3C' })

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides[0]!.background).toEqual({ type: 'solid', color: '#1A2B3C' })
  })

  it('replaces an existing bg without duplicating', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    setSlideBackground(slide, '#111111')
    setSlideBackground(slide, '#222222')
    expect((slide.bodyPrefix.match(/<p:bg>/g) ?? []).length).toBe(1)

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides[0]!.background).toEqual({ type: 'solid', color: '#222222' })
  })
})
