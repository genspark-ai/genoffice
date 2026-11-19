import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
const MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'

function anchorParagraph(wsp: string, drawingAttrs = ''): string {
  return (
    `<w:p><w:r><w:drawing${drawingAttrs}><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>584200</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>127000</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="584200" cy="374650"/><wp:wrapNone/>' +
    '<wp:docPr id="1" name="Shape 1"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `${wsp}</a:graphicData></a:graphic>` +
    '</wp:anchor></w:drawing></w:r></w:p>'
  )
}

describe('themeless documents (no word/theme/theme1.xml)', () => {
  it('resolves wps:style schemeClr against the built-in Office palette', async () => {
    // docx4j 2010-glow-then-Alt / 2010-mcAlternateCo: a gallery ellipse whose
    // only ink comes from fillRef/lnRef accent1 — colors:null dropped it
    const wsp =
      `<wps:wsp ${WPS_NS}><wps:cNvSpPr/><wps:spPr>` +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="584200" cy="374650"/></a:xfrm>' +
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr>' +
      '<wps:style>' +
      '<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>' +
      '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
      '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
      '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>' +
      '</wps:style><wps:bodyPr/></wps:wsp>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: anchorParagraph(wsp) }))
    const box = parsed.blocks.find((b) => b.textboxes?.length)?.textboxes?.[0]
    expect(box).toBeDefined()
    expect(box?.prst).toBe('ellipse')
    expect(box?.fill).toBe('4472C4')
    // accent1 at 50% shade
    expect(box?.borderColor).toBe('223962')
    expect(parsed.themeColors?.accent1).toBe('4472C4')
  })
})

describe('w:drawing with attributes (mc:MustUnderstand)', () => {
  it('extracts the shape from an attributed <w:drawing> open tag', async () => {
    const wsp =
      `<wps:wsp ${WPS_NS}><wps:cNvSpPr/><wps:spPr>` +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="2366645" cy="1404620"/></a:xfrm>' +
      '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
      '</wps:spPr><wps:bodyPr/></wps:wsp>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(wsp, ` ${MC_NS} mc:MustUnderstand="wps"`),
      }),
    )
    const box = parsed.blocks.find((b) => b.textboxes?.length)?.textboxes?.[0]
    expect(box).toBeDefined()
    expect(box?.fill).toBe('FF0000')
    expect(box?.floating).toBe(true)
  })
})

describe('external textbox part (wps:txbx r:txbx)', () => {
  it('reads the referenced word/txbx1.xml paragraphs into the box, display-only', async () => {
    // TestFiles mcdoc: the Choice textbox stores its content in a separate
    // part whose w14:txbx root carries the w:p list directly
    const wsp =
      `<wps:wsp ${WPS_NS}><wps:cNvSpPr txBox="1"/><wps:spPr>` +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="2366645" cy="1404620"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' +
      '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' +
      '</wps:spPr><wps:txbx r:txbx="rId20" txbxSeq="0"/><wps:bodyPr/></wps:wsp>'
    const txbxPart =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w14:txbx xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:r><w:t>hello</w:t></w:r></w:p></w14:txbx>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(wsp),
        extraRels:
          '<Relationship Id="rId20" Type="http://schemas.microsoft.com/office/2006/relationships/txbx" Target="txbx1.xml"/>',
        extraParts: [
          {
            path: 'word/txbx1.xml',
            xml: txbxPart,
            contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.txbx+xml',
          },
        ],
      }),
    )
    const box = parsed.blocks.find((b) => b.textboxes?.length)?.textboxes?.[0]
    expect(box).toBeDefined()
    expect(box?.fill).toBe('FFFFFF')
    expect(box?.borderColor).toBe('000000')
    expect(box?.paras.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['hello'])
    // the external content is not patchable in document.xml
    expect(box?.readOnly).toBe(true)
    expect(box?.txbxIndex).toBeUndefined()
  })
})
