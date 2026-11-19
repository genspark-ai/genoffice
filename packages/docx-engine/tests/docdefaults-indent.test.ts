/**
 * pPrDefault w:ind is the base every paragraph's indents inherit from (Google
 * Docs exports put the document's first-line indent there, with an empty
 * Normal). A centered in-line picture under it sits half the indent right of
 * center in Word and LibreOffice.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx, IMAGE_PARAGRAPH_XML } from './helpers/build-docx'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const stylesXml = (pPr: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
  `<w:docDefaults><w:rPrDefault/><w:pPrDefault><w:pPr>${pPr}</w:pPr></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="normal"/></w:style></w:styles>`

const CENTERED_PICTURE = IMAGE_PARAGRAPH_XML.replace(
  '<w:p>',
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
)

const ANCHORED_PICTURE =
  '<w:p><w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="p"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

describe('pPrDefault w:ind', () => {
  it('reaches docDefaults, a hanging indent as a negative first line', async () => {
    const first = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p/>',
        stylesXml: stylesXml('<w:ind w:left="120" w:right="60" w:firstLine="708"/>'),
      }),
    )
    expect(first.docDefaults).toMatchObject({
      indentLeftTwips: 120,
      indentRightTwips: 60,
      indentFirstLineTwips: 708,
    })
    const hanging = await parseDocx(
      await buildDocx({ bodyXml: '<w:p/>', stylesXml: stylesXml('<w:ind w:hanging="360"/>') }),
    )
    expect(hanging.docDefaults?.indentFirstLineTwips).toBe(-360)
  })

  it('a centered in-line picture without its own w:ind takes the default first line', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: CENTERED_PICTURE,
        stylesXml: stylesXml('<w:ind w:firstLine="708"/>'),
        withImage: true,
      }),
    )
    expect(doc.blocks[0].type).toBe('image')
    expect(doc.blocks[0].imageAlign).toBe('center')
    expect(doc.blocks[0].imageParagraphIndentFirstLine).toBe(708)
  })

  it("the picture paragraph's own w:ind still wins", async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: CENTERED_PICTURE.replace(
          '<w:jc w:val="center"/>',
          '<w:ind w:firstLine="0"/><w:jc w:val="center"/>',
        ),
        stylesXml: stylesXml('<w:ind w:firstLine="708"/>'),
        withImage: true,
      }),
    )
    expect(doc.blocks[0].imageParagraphIndentFirstLine).toBeUndefined()
  })

  it('a floating picture positions against the column and ignores it', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: ANCHORED_PICTURE,
        stylesXml: stylesXml('<w:ind w:firstLine="708"/>'),
        withImage: true,
      }),
    )
    expect(doc.blocks[0].imageWrap).toBe('square-left')
    expect(doc.blocks[0].imageParagraphIndentFirstLine).toBeUndefined()
  })
})
