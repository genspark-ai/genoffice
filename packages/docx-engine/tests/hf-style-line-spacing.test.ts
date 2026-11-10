/**
 * Word sizes header/footer lines by the strip paragraph's style chain: the
 * built-in Header/Footer styles reset Normal's w:line, so an 8pt header under a
 * 1.35-spaced Normal is a single line (Word probe 2026-09-04, SAS batch 2 098:
 * 15.5pt header, not 21pt). Direct w:spacing still wins.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const hdr = (pPr: string) =>
  `${XML_DECL}<w:hdr ${W}><w:p><w:pPr>${pPr}</w:pPr>` +
  '<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>Header line</w:t></w:r></w:p></w:hdr>'

async function headerPara(pPr: string) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraStylesXml:
      '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="TallHeader"><w:name w:val="tall header"/><w:basedOn w:val="Header"/>' +
      '<w:pPr><w:spacing w:line="360" w:lineRule="exact"/></w:pPr></w:style>',
    extraRels:
      '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    extraParts: [{ path: 'word/header1.xml', xml: hdr(pPr), contentType: HEADER_TYPE }],
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId60"/>',
  })
  return (await parseDocx(bytes)).headerParas![0]
}

describe('header/footer paragraph line spacing from the style chain', () => {
  it('takes the Header style single spacing when the paragraph declares none', async () => {
    expect(await headerPara('<w:pStyle w:val="Header"/>')).toMatchObject({
      lineRule: 'auto',
      lineRawTwips: 240,
      lineSpacing: 1,
    })
  })

  it('resolves through basedOn (exact line from the derived style)', async () => {
    const para = await headerPara('<w:pStyle w:val="TallHeader"/>')
    expect(para).toMatchObject({ lineRule: 'exact', lineRawTwips: 360 })
    expect(para.lineSpacing).toBeUndefined()
  })

  it('keeps direct w:spacing over the style', async () => {
    const para = await headerPara(
      '<w:pStyle w:val="Header"/><w:spacing w:line="480" w:lineRule="auto"/>',
    )
    expect(para).toMatchObject({ lineRule: 'auto', lineRawTwips: 480, lineSpacing: 2 })
  })

  it('leaves an unstyled paragraph without line spacing (inherits the document default)', async () => {
    const para = await headerPara('')
    expect(para.lineRule).toBeUndefined()
    expect(para.lineRawTwips).toBeUndefined()
  })
})
