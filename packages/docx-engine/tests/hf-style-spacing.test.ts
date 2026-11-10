/**
 * Word stacks header/footer paragraph spacing resolved through the style chain
 * and docDefaults like body paragraphs (corpus measurement 2026-09-17: a header
 * of a Calibri 11 blank line, a Normal (Web) blank line and another blank line
 * pushes the body 68.7pt below the header edge = 13.4 + 14 + 13.8 + 14 + 13.4).
 * Autospacing is Word's HTML auto 14pt (280 twips). A blank strip line is sized
 * by its mark, else by the style's run size.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const hdr = (paras: string) => `${XML_DECL}<w:hdr ${W}>${paras}</w:hdr>`
const p = (pPr: string, runs = '') => `<w:p><w:pPr>${pPr}</w:pPr>${runs}</w:p>`

async function headerParas(
  paras: string,
  docDefaultsPPr = '<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>',
  normalPPr = '',
) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    stylesXml:
      `${XML_DECL}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr>${docDefaultsPPr}</w:pPr></w:pPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${normalPPr}</w:style>` +
      '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="NormalWeb"><w:name w:val="Normal (Web)"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/></w:pPr>' +
      '<w:rPr><w:sz w:val="24"/></w:rPr></w:style></w:styles>',
    extraRels:
      '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    extraParts: [{ path: 'word/header1.xml', xml: hdr(paras), contentType: HEADER_TYPE }],
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId60"/>',
  })
  return (await parseDocx(bytes)).headerParas!
}

describe('header/footer paragraph spacing from the style chain', () => {
  it('a Normal (Web) blank line carries 14pt auto spacing above and below and its 12pt mark size', async () => {
    const [a, web, b] = await headerParas(
      p('<w:pStyle w:val="Header"/>') +
        p('<w:pStyle w:val="NormalWeb"/>') +
        p('<w:pStyle w:val="Header"/>'),
    )
    expect(a.spaceBefore).toBeUndefined()
    expect(a.spaceAfter).toBeUndefined()
    expect(web).toMatchObject({ spaceBefore: 280, spaceAfter: 280, emptyRunSizeHalfPoints: 24 })
    expect(b.spaceAfter).toBeUndefined()
  })

  it('an unstyled strip paragraph takes the docDefaults spacing', async () => {
    const [para] = await headerParas(p(''))
    expect(para.spaceAfter).toBe(200)
    expect(para.spaceBefore).toBeUndefined()
  })

  it('... unless the default paragraph style overrides it (Normal after=0 header lines stay tight)', async () => {
    const [para] = await headerParas(
      p(''),
      '<w:spacing w:after="160" w:line="278" w:lineRule="auto"/>',
      '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>',
    )
    expect(para.spaceAfter).toBeUndefined()
  })

  it('direct w:spacing wins, and a direct autospacing off restores the literal', async () => {
    const [direct, autoOff] = await headerParas(
      p('<w:pStyle w:val="NormalWeb"/><w:spacing w:before="60" w:after="0"/>') +
        p(
          '<w:pStyle w:val="NormalWeb"/><w:spacing w:before="60" w:beforeAutospacing="0" w:after="40" w:afterAutospacing="0"/>',
        ),
    )
    // the style's autospacing still applies to the direct literals (Word resolves the flags independently)
    expect(direct).toMatchObject({ spaceBefore: 280, spaceAfter: 280 })
    expect(autoOff).toMatchObject({ spaceBefore: 60, spaceAfter: 40 })
  })

  it('the paragraph mark size beats the style size for a blank line; text runs get none', async () => {
    const [mark, text] = await headerParas(
      p('<w:pStyle w:val="NormalWeb"/><w:rPr><w:sz w:val="16"/></w:rPr>') +
        p('<w:pStyle w:val="NormalWeb"/>', '<w:r><w:t>x</w:t></w:r>'),
    )
    expect(mark.emptyRunSizeHalfPoints).toBe(16)
    expect(text.emptyRunSizeHalfPoints).toBeUndefined()
  })
})
