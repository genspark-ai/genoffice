import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** LibreOffice writes the field begin marker as an empty element pair
 *  (<w:fldChar ...></w:fldChar>) instead of a self-closing tag */
const HEADER =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> SUBJECT </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>Annual Report</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:tab/><w:tab/><w:t xml:space="preserve">Page </w:t></w:r>' +
  '<w:r><w:rPr><w:rStyle w:val="PageNumber"/></w:rPr><w:fldChar w:fldCharType="begin"></w:fldChar></w:r>' +
  '<w:r><w:rPr><w:rStyle w:val="PageNumber"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  '<w:r><w:rPr><w:rStyle w:val="PageNumber"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:rPr><w:rStyle w:val="PageNumber"/></w:rPr><w:t>35</w:t></w:r>' +
  '<w:r><w:rPr><w:rStyle w:val="PageNumber"/></w:rPr><w:fldChar w:fldCharType="end"></w:fldChar></w:r>' +
  '</w:p></w:hdr>'

describe('header fields with element-form fldChar markers', () => {
  it('substitutes the live PAGE mark and keeps the cached SUBJECT result', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="HeaderBase"><w:name w:val="Header base"/><w:basedOn w:val="Normal"/>' +
        '<w:pPr><w:tabs><w:tab w:val="clear" w:pos="449"/><w:tab w:val="center" w:pos="4819"/>' +
        '<w:tab w:val="right" w:pos="9638"/></w:tabs></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="HeaderBase"/>' +
        '<w:pPr><w:suppressLineNumbers/></w:pPr></w:style>',
      extraRels:
        '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: HEADER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId61"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.headerHasPageNumber).toBe(true)
    expect(parsed.headerText).toBe(`Annual ReportPage ${PAGE_MARK}`)
    const para = parsed.headerParas![0]
    expect(para.runs.map((r) => r.text).join('')).toBe(`Annual Report\t\tPage ${PAGE_MARK}`)
    expect(para.tabStops).toEqual([
      { pos: 4819, val: 'center' },
      { pos: 9638, val: 'right' },
    ])
  })
})
