import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const FOOTER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const ftrPart = (inner: string) => `${XML_DECL}<w:ftr ${W_NS}>${inner}</w:ftr>`

describe('blank header/footer parts', () => {
  it('keeps the empty paragraphs of an all-blank footer with their spacing', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId80" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
      extraParts: [
        {
          path: 'word/footer1.xml',
          xml: ftrPart(
            '<w:p><w:pPr><w:spacing w:after="200" w:line="288" w:lineRule="auto"/></w:pPr></w:p>' +
              '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>',
          ),
          contentType: FOOTER_TYPE,
        },
      ],
      sectPrExtra: '<w:footerReference w:type="default" r:id="rId80"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.footerText).toBe('')
    expect(parsed.footerParas).toHaveLength(2)
    expect(parsed.footerParas![0]).toMatchObject({ spaceAfter: 200, lineRawTwips: 288 })
    expect(parsed.footerParas![0].runs).toEqual([])
    expect(parsed.footerParas![1].runs).toEqual([])
  })
})
