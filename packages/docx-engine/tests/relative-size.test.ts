import { describe, expect, it } from 'vitest'
import { patchTextboxSizes } from '../src/generate'
import {
  drawingRelSizeOf,
  parseDocx,
  relSizeTwips,
  type SectionSettings,
  type TextboxDisplay,
} from '../src/index'
import { setDrawingRelSize, setVmlRelSize } from '../src/relative-size'
import { buildDocx } from './helpers/build-docx'

const rel = (axis: 'H' | 'V', from: string, units: number) => {
  const pct = axis === 'H' ? 'pctWidth' : 'pctHeight'
  return `<wp14:sizeRel${axis} relativeFrom="${from}"><wp14:${pct}>${units}</wp14:${pct}></wp14:sizeRel${axis}>`
}

const shape = (tail: string, bodyPr = '<wps:bodyPr><a:noAutofit/></wps:bodyPr>') =>
  '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="2743200" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Box"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
  `${bodyPr}</wps:wsp></a:graphicData></a:graphic>${tail}</wp:anchor></w:drawing></mc:Choice>` +
  '<mc:Fallback><w:pict><v:shape style="position:absolute;width:216pt;height:72pt;mso-width-percent:500;mso-width-relative:margin">' +
  '<v:textbox><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>' +
  '</mc:AlternateContent></w:r></w:p>'

const boxesOf = async (bodyXml: string): Promise<TextboxDisplay[]> =>
  (await parseDocx(await buildDocx({ bodyXml }))).blocks[0].textboxes ?? []

describe('wp14:sizeRelH / sizeRelV', () => {
  it('is read as a share of what it names; a zero share is an absolute side', async () => {
    const tail = rel('H', 'margin', 50000) + rel('V', 'page', 0)
    expect(drawingRelSizeOf(shape(tail))).toEqual({ w: { from: 'margin', pct: 50 } })
    const [box] = await boxesOf(shape(tail))
    expect(box.sizeRel).toEqual({ w: { from: 'margin', pct: 50 } })
    // the cached extent is what the page draws
    expect(box.widthPx).toBe(288)
    expect(
      (await boxesOf(shape(rel('H', 'margin', 0) + rel('V', 'margin', 0))))[0].sizeRel,
    ).toBeUndefined()
  })

  it('works out to a share of the page, the text area or a margin', () => {
    const section = {
      pageWidth: 12240,
      pageHeight: 15840,
      marginTop: 1440,
      marginRight: 1800,
      marginBottom: 720,
      marginLeft: 1080,
    } as SectionSettings
    expect(relSizeTwips(section, 'H', { from: 'margin', pct: 50 })).toBe(4680)
    expect(relSizeTwips(section, 'H', { from: 'page', pct: 25 })).toBe(3060)
    expect(relSizeTwips(section, 'H', { from: 'rightMargin', pct: 100 })).toBe(1800)
    expect(relSizeTwips(section, 'V', { from: 'margin', pct: 10 })).toBe(1368)
    expect(relSizeTwips(section, 'V', { from: 'bottomMargin', pct: 50 })).toBe(360)
  })

  it('a resize alone makes the side absolute and leaves the other as the file has it', () => {
    const xml = shape(rel('H', 'margin', 50000) + rel('V', 'page', 20000))
    const out = patchTextboxSizes(xml, [{ wPx: 200, hPx: null }])
    expect(out).toContain(rel('H', 'margin', 0))
    expect(out).toContain(rel('V', 'page', 20000))
    expect(out).toContain('<wp:extent cx="1905000" cy="914400"/>')
    expect(out).not.toContain('mso-width-percent')
  })

  it('writes a share with the cached size beside it, into a drawing that had none too', () => {
    const had = patchTextboxSizes(shape(rel('H', 'margin', 0) + rel('V', 'margin', 0)), [
      { wPx: 312, relW: { from: 'page', pct: 38.2 } },
    ])
    expect(had).toContain(rel('H', 'page', 38200))
    expect(had).toContain('<wp:extent cx="2971800" cy="914400"/>')
    expect(had).toContain('mso-width-percent:382;mso-width-relative:page')
    const bare = patchTextboxSizes(shape(''), [{ hPx: 96, relH: { from: 'topMargin', pct: 100 } }])
    expect(bare).toMatch(
      /<\/a:graphic><wp14:sizeRelV xmlns:wp14="[^"]+" relativeFrom="topMargin"><wp14:pctHeight>100000<\/wp14:pctHeight><\/wp14:sizeRelV><\/wp:anchor>/,
    )
    expect(bare).toContain('mso-height-percent:1000;mso-height-relative:top-margin-area')
  })

  it('leaves a drawing nobody resized byte for byte', () => {
    const xml = shape(rel('H', 'margin', 50000) + rel('V', 'page', 20000))
    expect(patchTextboxSizes(xml, [null])).toBe(xml)
    expect(setDrawingRelSize(shape(''), 'H', null)).toBe(shape(''))
  })

  it('puts a new width before an existing height', () => {
    const out = setDrawingRelSize(shape(rel('V', 'page', 0)), 'H', { from: 'margin', pct: 40 })
    expect(out).toMatch(
      /<wp14:sizeRelH [^>]*relativeFrom="margin">.*<wp14:sizeRelV relativeFrom="page">/,
    )
    expect(setVmlRelSize('width:10pt;mso-width-percent:400', 'H', null)).toBe('width:10pt')
  })
})
