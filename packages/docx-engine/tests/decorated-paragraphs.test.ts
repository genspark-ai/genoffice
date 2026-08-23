import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

/**
 * Documents produced by design tools (and Word itself) pair every DrawingML
 * shape with a legacy VML fallback:
 *   <mc:AlternateContent><mc:Choice><w:drawing>…</mc:Choice>
 *   <mc:Fallback><w:pict>…</mc:Fallback></mc:AlternateContent>
 * The raw bytes therefore contain <w:pict> even for a plain decorated text
 * paragraph — classification must ignore fallback content or whole documents
 * degrade into embedded-object chips.
 */

const THIN_RULE_DRAWING =
  '<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor behindDoc="1" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="6048375" cy="9525"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6048375" cy="9525"/></a:xfrm></wps:spPr>' +
  '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
  '<mc:Fallback><w:pict>' +
  '<v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute" id="rule1" filled="true"/>' +
  '</w:pict></mc:Fallback></mc:AlternateContent></w:r>'

describe('paragraphs with decorative shapes + VML fallbacks', () => {
  it('keeps text of a decorated heading editable (fallback w:pict must not trigger Embedded object)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${THIN_RULE_DRAWING}` +
          '<w:r><w:t>Mandatory Story Header Format</w:t></w:r></w:p>',
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('heading')
    expect(block.level).toBe(2)
    expect(block.runs?.map((r) => r.text).join('')).toBe('Mandatory Story Header Format')
  })

  it('keeps decorated list items as listItem', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
          `${THIN_RULE_DRAWING}<w:r><w:t>Include new date only on a time skip</w:t></w:r></w:p>`,
        withNumbering: true,
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('listItem')
    expect(block.runs?.map((r) => r.text).join('')).toBe('Include new date only on a time skip')
  })

  it('marks a text-less thin shape as a decorative rule', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: `<w:p>${THIN_RULE_DRAWING}</w:p>` }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Drawing object')
    expect(block.decorative).toBe(true)
  })

  it('still protects real VML embedded objects (w:pict outside mc:Fallback)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:pict>' +
          '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="ole1" style="width:100pt;height:100pt"/>' +
          '</w:pict></w:r></w:p>',
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Embedded object')
  })

  it('extracts anchored textbox content as a display-only model (code box)', async () => {
    const textboxDrawing =
      '<w:r><w:drawing><wp:anchor behindDoc="1" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:extent cx="6038850" cy="1704975"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>' +
      '<a:ln w="9524"><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></wps:spPr>' +
      '<wps:txbx><w:txbxContent>' +
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Courier New"/><w:color w:val="415461"/><w:sz w:val="22"/></w:rPr>' +
      '<w:t>**Current date:** {Month}</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>CRITICAL</w:t></w:r></w:p>' +
      '</w:txbxContent></wps:txbx></wps:wsp>' +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    const doc = await parseDocx(await buildDocx({ bodyXml: `<w:p>${textboxDrawing}</w:p>` }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Text box')
    const box = block.textboxes?.[0]
    expect(box?.fill).toBe('F5F5F5')
    expect(box?.borderColor).toBe('CCCCCC')
    expect(box?.paras).toHaveLength(2)
    expect(box?.paras[0].runs[0]).toMatchObject({
      text: '**Current date:** {Month}',
      color: '415461',
      font: 'Courier New',
    })
    expect(box?.paras[1].align).toBe('center')
    expect(box?.paras[1].runs[0]).toMatchObject({ text: 'CRITICAL', bold: true })
  })

  it('preserves the original XML slice for decorated paragraphs (byte-level patch anchor)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p>${THIN_RULE_DRAWING}<w:r><w:t>text</w:t></w:r></w:p>`,
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    expect(block.originalXml).toContain('<mc:Fallback>')
    expect(block.originalXml).toContain('<w:pict>')
  })
})

describe('picture content controls (w:picture in sdtPr, tdf#165359)', () => {
  it('an sdt-wrapped inline picture with text stays an editable paragraph, not an Embedded object chip', async () => {
    const drawing =
      '<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:sdt><w:sdtPr><w:picture/></w:sdtPr>' +
          `<w:sdtContent><w:r>${drawing}</w:r></w:sdtContent></w:sdt>` +
          '<w:r><w:t>ITEM NUMBER</w:t></w:r></w:p>',
        withImage: true,
      }),
    )
    const block = doc.blocks[0]
    expect(block.label).not.toBe('Embedded object')
    expect(JSON.stringify(block.runs?.map((r) => r.text))).toContain('ITEM NUMBER')
    expect(block.runs?.some((r) => r.image)).toBe(true)
  })
})
