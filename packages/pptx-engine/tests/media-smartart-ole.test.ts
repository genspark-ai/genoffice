/**
 * Parsing tests for media (audio/video) / SmartArt pre-render / OLE preview + graphicFrame p:xfrm patch.
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { openPptx, createBlankPptx } from '../src/index'
import { parseSlide } from '../src/parse'
import { patchElementXfrm } from '../src/generate'
import type { PassthroughElement, PictureElement, TextElement } from '../src/types'

const theme: any = { colors: { accent1: '#4472C4', lt1: '#FFFFFF', dk1: '#000000' } }

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

describe('audio/video pictures (p:nvPr a:videoFile/a:audioFile)', () => {
  const pic = (av: string) =>
    '<p:pic><p:nvPicPr><p:cNvPr id="4" name="movie.mp4"/><p:cNvPicPr/>' +
    `<p:nvPr>${av}</p:nvPr></p:nvPicPr>` +
    '<p:blipFill><a:blip r:embed="rId3"/></p:blipFill>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>'

  it('video: kind + embedded media path parsed, poster frame still goes through mediaRef', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(pic('<a:videoFile r:link="rId2"/>')),
      ctx: {
        theme,
        mediaRels: new Map([['rId3', 'ppt/media/poster1.png']]),
        avRels: new Map([['rId2', { target: 'ppt/media/media1.mp4' }]]),
      },
    })
    const el = slide.elements[0] as PictureElement
    expect(el.type).toBe('picture')
    expect(el.mediaRef).toBe('ppt/media/poster1.png')
    expect(el.media).toEqual({ kind: 'video', target: 'ppt/media/media1.mp4' })
  })

  it('audio: external URL keeps the external flag', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(pic('<a:audioFile r:link="rId2"/>')),
      ctx: {
        theme,
        avRels: new Map([['rId2', { target: 'https://example.com/a.mp3', external: true }]]),
      },
    })
    const el = slide.elements[0] as PictureElement
    expect(el.media).toEqual({ kind: 'audio', target: 'https://example.com/a.mp3', external: true })
  })

  it('plain pictures carry no media field', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(pic('')),
      ctx: { theme },
    })
    expect((slide.elements[0] as PictureElement).media).toBeUndefined()
  })
})

const SMARTART_FRAME =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Diagram 1"/></p:nvGraphicFramePr>' +
  '<p:xfrm><a:off x="100" y="200"/><a:ext cx="4000" cy="3000"/></p:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
  '<dgm:relIds xmlns:dgm="dgm" r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>' +
  '</a:graphicData></a:graphic></p:graphicFrame>'

const DIAGRAM_DRAWING =
  '<dsp:drawing xmlns:dsp="dsp" xmlns:a="a"><dsp:spTree><dsp:nvGrpSpPr/><dsp:grpSpPr/>' +
  '<dsp:sp modelId="{X}"><dsp:nvSpPr><dsp:cNvPr id="1" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>' +
  '<dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/></a:xfrm>' +
  '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></dsp:spPr>' +
  '<dsp:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>' +
  '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"/>' +
  '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>' +
  '<dsp:txBody><a:bodyPr/><a:p><a:r><a:t>Step One</a:t></a:r></a:p></dsp:txBody></dsp:sp>' +
  '</dsp:spTree></dsp:drawing>'

describe('SmartArt prerendered drawing read-only preview', () => {
  it('parses previewShapes: shape colors from dsp:style theme references, text color from fontRef', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(SMARTART_FRAME),
      ctx: { theme, diagramDrawings: new Map([['rId1', DIAGRAM_DRAWING]]) },
    })
    const el = slide.elements[0] as PassthroughElement
    expect(el.type).toBe('passthrough')
    expect(el.kind).toBe('smartart')
    expect(el.transform.offset).toEqual({ x: 100, y: 200, cx: 4000, cy: 3000 })
    expect(el.previewShapes?.length).toBe(1)
    const sp = el.previewShapes![0] as TextElement
    expect(sp.presetGeometry).toBe('roundRect')
    expect(sp.fill).toEqual({ type: 'solid', color: '#4472C4' })
    expect(sp.stroke?.fill).toEqual({ type: 'solid', color: '#4472C4' })
    const run = sp.text!.paragraphs[0]!.runs[0]!
    expect(run.text).toBe('Step One')
    expect(run.color).toBe('#FFFFFF')
  })

  it('without a drawing part (ctx missing) it remains a plain placeholder passthrough', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(SMARTART_FRAME),
      ctx: { theme },
    })
    const el = slide.elements[0] as PassthroughElement
    expect(el.kind).toBe('smartart')
    expect(el.previewShapes).toBeUndefined()
  })

  it('save byte fidelity: anchor.originalXml not polluted by preview parsing', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(SMARTART_FRAME),
      ctx: { theme, diagramDrawings: new Map([['rId1', DIAGRAM_DRAWING]]) },
    })
    expect(slide.elements[0]!.anchor.originalXml).toBe(SMARTART_FRAME)
  })

  it('dsp:txXfrm text frame keeps fontRef color but not the shape effectRef', () => {
    const drawing =
      '<dsp:drawing xmlns:dsp="dsp" xmlns:a="a"><dsp:spTree><dsp:nvGrpSpPr/><dsp:grpSpPr/>' +
      '<dsp:sp modelId="{X}"><dsp:nvSpPr><dsp:cNvPr id="1" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>' +
      '<dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/></a:xfrm>' +
      '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></dsp:spPr>' +
      '<dsp:txXfrm><a:off x="100" y="200"/><a:ext cx="1800" cy="600"/></dsp:txXfrm>' +
      '<dsp:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>' +
      '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
      '<a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef>' +
      '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>' +
      '<dsp:txBody><a:bodyPr/><a:p><a:r><a:t>Step One</a:t></a:r></a:p></dsp:txBody></dsp:sp>' +
      '</dsp:spTree></dsp:drawing>'
    const effectTheme: any = {
      colors: { accent1: '#4472C4', lt1: '#FFFFFF', dk1: '#000000' },
      effectStyles: [
        { 'a:effectStyle': { 'a:effectLst': {} } },
        { 'a:effectStyle': { 'a:effectLst': {} } },
        {
          'a:effectStyle': {
            'a:effectLst': {
              'a:glow': {
                '@_rad': '63500',
                'a:schemeClr': { '@_val': 'phClr', 'a:alpha': { '@_val': '40000' } },
              },
              'a:outerShdw': {
                '@_blurRad': '57150',
                '@_dist': '19050',
                '@_dir': '5400000',
                'a:srgbClr': { '@_val': '000000', 'a:alpha': { '@_val': '63000' } },
              },
            },
          },
        },
      ],
    }
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(SMARTART_FRAME),
      ctx: { theme: effectTheme, diagramDrawings: new Map([['rId1', drawing]]) },
    })
    const shapes = (slide.elements[0] as PassthroughElement).previewShapes as TextElement[]
    expect(shapes).toHaveLength(2)
    const [shape, text] = shapes
    expect(shape!.presetGeometry).toBe('roundRect')
    expect(shape!.shadow).toBeTruthy()
    expect(shape!.glow).toBeTruthy()
    expect(shape!.text).toBeUndefined()
    expect(text!.fill).toEqual({ type: 'none' })
    expect(text!.stroke).toBeUndefined()
    expect(text!.shadow).toBeUndefined()
    expect(text!.glow).toBeUndefined()
    expect(text!.text!.paragraphs[0]!.runs[0]!.text).toBe('Step One')
    expect(text!.text!.paragraphs[0]!.runs[0]!.color).toBe('#FFFFFF')
  })
})

const OLE_FRAME =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Object 1"/></p:nvGraphicFramePr>' +
  '<p:xfrm><a:off x="10" y="20"/><a:ext cx="900" cy="600"/></p:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">' +
  '<mc:AlternateContent xmlns:mc="mc"><mc:Choice Requires="v"><p:oleObj spid="_x0000_s1026"/></mc:Choice>' +
  '<mc:Fallback><p:oleObj><p:pic><p:nvPicPr><p:cNvPr id="7" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rId7"/></p:blipFill>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="900" cy="600"/></a:xfrm></p:spPr></p:pic></p:oleObj></mc:Fallback>' +
  '</mc:AlternateContent></a:graphicData></a:graphic></p:graphicFrame>'

describe('OLE embedded preview picture', () => {
  it('drills through mc:AlternateContent to use the fallback p:pic as preview', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(OLE_FRAME),
      ctx: { theme, mediaRels: new Map([['rId7', 'ppt/media/image9.emf']]) },
    })
    const el = slide.elements[0] as PassthroughElement
    expect(el.kind).toBe('ole')
    expect(el.previewPicture?.mediaRef).toBe('ppt/media/image9.emf')
  })
})

describe('graphicFrame p:xfrm patch (table/chart/smartart/ole movable)', () => {
  it('replaces the p:xfrm off/ext without touching the embedded pic a:xfrm', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(OLE_FRAME),
      ctx: { theme },
    })
    const el = slide.elements[0]!
    el.transform = { ...el.transform, offset: { x: 111, y: 222, cx: 900, cy: 600 } }
    const patched = patchElementXfrm(el, el.anchor.originalXml)
    expect(patched).toContain('<p:xfrm><a:off x="111" y="222"/><a:ext cx="900" cy="600"/></p:xfrm>')
    // The embedded preview picture's own a:xfrm kept as-is
    expect(patched).toContain('<a:xfrm><a:off x="0" y="0"/><a:ext cx="900" cy="600"/></a:xfrm>')
  })
})

describe('openPptx end-to-end: slide rels → diagramData → dataModelExt → drawing part / video rels', () => {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

  it('assembles ctx from the zip: SmartArt previewShapes + video media parsed end-to-end', async () => {
    const bytes = await createBlankPptx()
    const zip = await JSZip.loadAsync(bytes)

    const videoPic =
      '<p:pic><p:nvPicPr><p:cNvPr id="20" name="movie.mp4"/><p:cNvPicPr/>' +
      '<p:nvPr><a:videoFile r:link="rId12"/></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId13"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500" cy="400"/></a:xfrm></p:spPr></p:pic>'
    const smartartFrame = SMARTART_FRAME.replace('r:dm="rId1"', 'r:dm="rId10"')

    const slideXml = String(await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
      '</p:spTree>',
      `${smartartFrame}${videoPic}</p:spTree>`,
    )
    zip.file('ppt/slides/slide1.xml', slideXml)

    const slideRels = String(
      await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string'),
    ).replace(
      '</Relationships>',
      `<Relationship Id="rId10" Type="${R}/diagramData" Target="../diagrams/data1.xml"/>` +
        `<Relationship Id="rId11" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>` +
        `<Relationship Id="rId12" Type="${R}/video" Target="../media/media1.mp4"/>` +
        `<Relationship Id="rId13" Type="${R}/image" Target="../media/poster1.png"/>` +
        '</Relationships>',
    )
    zip.file('ppt/slides/_rels/slide1.xml.rels', slideRels)

    zip.file(
      'ppt/diagrams/data1.xml',
      '<?xml version="1.0"?><dgm:dataModel xmlns:dgm="dgm" xmlns:dsp="dsp">' +
        '<dgm:extLst><a:ext xmlns:a="a" uri="{X}">' +
        '<dsp:dataModelExt relId="rId11" minVer="v"/></a:ext></dgm:extLst></dgm:dataModel>',
    )
    zip.file('ppt/diagrams/drawing1.xml', DIAGRAM_DRAWING)
    zip.file('ppt/media/media1.mp4', new Uint8Array([0]))
    zip.file('ppt/media/poster1.png', new Uint8Array([0]))

    const rebuilt = await zip.generateAsync({ type: 'uint8array' })
    const { deck } = await openPptx(rebuilt)
    const els = deck.slides[0]!.elements

    const sa = els.find((e) => e.type === 'passthrough') as PassthroughElement
    expect(sa.kind).toBe('smartart')
    expect(sa.previewShapes?.length).toBe(1)
    expect((sa.previewShapes![0] as TextElement).presetGeometry).toBe('roundRect')

    const vid = els.find((e) => e.type === 'picture') as PictureElement
    expect(vid.mediaRef).toBe('ppt/media/poster1.png')
    expect(vid.media).toEqual({ kind: 'video', target: 'ppt/media/media1.mp4' })
  })
})

describe('sp a:xfrm patch regression', () => {
  it('sp elements still take the a:xfrm path', () => {
    const sp =
      '<p:sp><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="10" cy="10"/></a:xfrm>' +
      '<a:prstGeom prst="rect"/></p:spPr></p:sp>'
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(sp),
      ctx: { theme },
    })
    const el = slide.elements[0]!
    el.transform = { ...el.transform, offset: { x: 5, y: 6, cx: 10, cy: 10 } }
    const patched = patchElementXfrm(el, el.anchor.originalXml)
    expect(patched).toContain('<a:off x="5" y="6"/>')
  })
})
