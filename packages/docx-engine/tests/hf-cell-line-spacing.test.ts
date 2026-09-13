import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const HEADER_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:pPr><w:pStyle w:val="HdrCell"/></w:pPr><w:r><w:t>University</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr><w:r><w:t>Author</w:t></w:r></w:p></w:tc>' +
  '</w:tr></w:tbl><w:p/></w:hdr>'

// Normal is double-spaced; the cell style resets it to single (Word sizes the
// row by the cell paragraph's own style chain)
const STYLES =
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="HdrCell"><w:name w:val="Hdr Cell"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:style>'

describe('header table cell paragraph line spacing', () => {
  it('carries the style-chain or direct w:spacing line into the cell paragraph props', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
        extraStylesXml: STYLES,
        extraRels:
          '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
        sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
        extraParts: [
          {
            path: 'word/header1.xml',
            xml: HEADER_XML,
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
          },
        ],
      }),
    )
    const row = Object.values(doc.hfParts ?? {})[0]?.paras[0]
    expect(row?.cells).toHaveLength(2)
    expect(row?.cells?.[0].paraProps?.[0]).toMatchObject({
      lineRule: 'auto',
      lineRawTwips: 240,
      lineSpacing: 1,
    })
    expect(row?.cells?.[1].paraProps?.[0]).toMatchObject({
      lineRule: 'auto',
      lineRawTwips: 480,
      lineSpacing: 2,
    })
  })
})
