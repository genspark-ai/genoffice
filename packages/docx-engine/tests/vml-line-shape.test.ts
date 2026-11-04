import { describe, expect, it } from 'vitest'
import type { Block } from '../src/types'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

/**
 * Issue #2481: a contract's blanks are legacy VML lines. Word pairs every wps
 * connector (mc:Choice) with a <v:line> (mc:Fallback): the Choice rendering has
 * no shape a text paragraph can show (flat connector, wrapNone), so the VML
 * line is the only drawable geometry. Word draws it as a long rule floating
 * over the blank; these tests pin the parse model the renderer needs.
 */

const NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office"'

interface ConnectorOptions {
  id: number
  xEmu: number
  yEmu: number
  widthEmu: number
  /** style tail of the VML twin: the blank's box + margin offsets */
  vmlStyle: string
}

/** one mc:AlternateContent run: wps connector Choice + VML <v:line> Fallback */
function connectorRun(opts: ConnectorOptions): string {
  const anchor =
    `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="${opts.id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${opts.xEmu}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${opts.yEmu}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${opts.widthEmu}" cy="0"/><wp:effectExtent l="0" t="0" r="19050" b="19050"/>` +
    `<wp:wrapNone/><wp:docPr id="${opts.id}" name="Connector ${opts.id}"/>` +
    `<wp:cNvGraphicFramePr/><a:graphic>` +
    `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvCnPr/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${opts.widthEmu}" cy="0"/></a:xfrm>` +
    `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>` +
    `<a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>` +
    `</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor>`
  const vmlLine =
    `<v:line id="_x0000_s1${opts.id}" o:spt="20" style="position:absolute;left:0pt;${opts.vmlStyle}" ` +
    `filled="f" stroked="t" coordsize="21600,21600"><v:fill on="f" focussize="0,0"/>` +
    `<v:stroke weight="0.5pt" color="#000000 [3213]" miterlimit="8" joinstyle="miter"/>` +
    `<v:imagedata o:title=""/><o:lock v:ext="edit" aspectratio="f"/></v:line>`
  return (
    `<w:r><mc:AlternateContent ${NS}>` +
    `<mc:Choice Requires="wps"><w:drawing>${anchor}</w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict>${vmlLine}</w:pict></mc:Fallback>` +
    `</mc:AlternateContent></w:r>`
  )
}

const textBlock = (): string =>
  '<w:p><w:r><w:rPr><w:sz w:val="27"/></w:rPr><w:t xml:space="preserve">甲方：           </w:t></w:r>' +
  connectorRun({
    id: 1,
    xEmu: 467360,
    yEmu: 363220,
    widthEmu: 2239010,
    vmlStyle:
      'margin-left:36.8pt;margin-top:28.6pt;height:0pt;width:176.35pt;z-index:251661312;' +
      'mso-width-relative:page;mso-height-relative:page;',
  }) +
  '<w:r><w:rPr><w:sz w:val="27"/></w:rPr><w:t xml:space="preserve">以下简称</w:t></w:r></w:p>'

function imageOf(block: Block): NonNullable<NonNullable<Block['runs']>[number]['image']> {
  const run = (block.runs ?? []).find((r) => r.image !== undefined)
  expect(run, 'a run carrying the VML line image').toBeTruthy()
  return run!.image!
}

describe('VML v:line blanks (issue #2481)', () => {
  it('parses a text paragraph blank as a floating horizontal rule image', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: textBlock() }))
    const block = doc.blocks[0]!
    // the paragraph stays editable: the line is a display-only run image
    expect(block.type).toBe('paragraph')
    expect((block.runs ?? []).map((r) => r.text).join('')).toBe('甲方：           以下简称')
    const image = imageOf(block)
    expect(image).toMatchObject({
      widthPx: 235,
      heightPx: 2,
      wrap: 'front',
      offsetXEmu: 467360,
      offsetYEmu: 363220,
    })
    // the pict fragment re-emits verbatim when the paragraph regenerates
    expect(image.xml).toContain('<v:line')
    expect(image.xml).toContain('width:176.35pt')
    const svg = decodeURIComponent(image.dataUrl.replace('data:image/svg+xml,', ''))
    expect(svg).toContain('<line x1="0" y1="1" x2="235" y2="1"')
    expect(svg).toContain('stroke="#000000"')
    expect(svg).toContain('stroke-width="1"')
  })

  it('parses an empty anchor paragraph blank as an anchored decorative rule', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p>' +
          connectorRun({
            id: 2,
            xEmu: 467360,
            yEmu: 3175,
            widthEmu: 2239010,
            vmlStyle:
              'margin-left:36.8pt;margin-top:0.25pt;height:0pt;width:176.3pt;z-index:251662336;',
          }) +
          '</w:p>',
      }),
    )
    const block = doc.blocks[0]!
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Drawing object')
    expect(block.decorative).toBe(true)
    // the rule keeps its own empty line and lands at the connector's offsets
    expect(block.anchorLine).toBeTruthy()
    expect(block.ruleWidthPx).toBe(235)
    expect(block.imageOffsetXEmu).toBe(467360)
    expect(block.imageOffsetYEmu).toBe(3175)
  })

  it('draws a direct vertical v:line pict as a tall thin rule image', async () => {
    const bodyXml =
      '<w:p><w:r><w:t>签署人：</w:t></w:r><w:r><w:pict>' +
      '<v:line xmlns:v="urn:schemas-microsoft-com:vml" id="_x0000_s3" o:spt="20" ' +
      'style="position:absolute;left:0pt;margin-left:10pt;margin-top:5pt;height:120pt;width:0pt" ' +
      'stroked="t"><v:stroke weight="0.5pt" color="#000000"/></v:line>' +
      '</w:pict></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    const block = doc.blocks[0]!
    expect(block.type).toBe('paragraph')
    const image = imageOf(block)
    expect(image).toMatchObject({
      widthPx: 2,
      heightPx: 160,
      wrap: 'front',
      offsetXEmu: 127000,
      offsetYEmu: 63500,
    })
    const svg = decodeURIComponent(image.dataUrl.replace('data:image/svg+xml,', ''))
    // a vertical stroke: both x coordinates share the middle of the 2px box
    expect(svg).toContain('<line x1="1" y1="0" x2="1" y2="160"')
  })
})
