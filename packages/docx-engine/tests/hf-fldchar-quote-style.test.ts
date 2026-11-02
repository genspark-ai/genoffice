import { describe, expect, it } from 'vitest'
import { PAGE_MARK, TOTAL_PAGES_MARK, parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const FTR_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
const HDR_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

const hdrPart = (inner: string) => `${XML_DECL}<w:hdr ${W_NS}>${inner}</w:hdr>`
const ftrPart = (inner: string) => `${XML_DECL}<w:ftr ${W_NS}>${inner}</w:ftr>`

/** complex field with the quote style of every attribute left to the caller */
const complexField = (q: '"' | "'", instr: string, cached: string) =>
  `<w:r><w:fldChar w:fldCharType=${q}begin${q}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType=${q}separate${q}/></w:r>` +
  `<w:r><w:t>${cached}</w:t></w:r>` +
  `<w:r><w:fldChar w:fldCharType=${q}end${q}/></w:r>`

const simpleField = (q: '"' | "'", instr: string, cached: string) =>
  `<w:fldSimple w:instr=${q} ${instr} ${q}>` + `<w:r><w:t>${cached}</w:t></w:r>` + `</w:fldSimple>`

async function parseFooter(inner: string) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    extraParts: [{ path: 'word/footer1.xml', xml: ftrPart(inner), contentType: FTR_TYPE }],
    sectPrExtra: '<w:footerReference w:type="default" r:id="rId90"/>',
  })
  return parseDocx(bytes)
}

describe('single-quoted header/footer field markers', () => {
  it('substitutes the live PAGE mark for a single-quoted fldChar PAGE field', async () => {
    const parsed = await parseFooter(`<w:p>${complexField("'", 'PAGE', '35')}</w:p>`)
    expect(parsed.footerText).toBe(PAGE_MARK)
    expect(parsed.footerHasPageNumber).toBe(true)
    expect(parsed.footerParas![0].runs.map((r) => r.text).join('')).toBe(PAGE_MARK)
  })

  it('substitutes the total-page mark for a single-quoted fldChar NUMPAGES field', async () => {
    const parsed = await parseFooter(`<w:p>${complexField("'", 'NUMPAGES', '12')}</w:p>`)
    expect(parsed.footerText).toBe(TOTAL_PAGES_MARK)
  })

  it('substitutes both marks for a single-quoted fldSimple PAGE / NUMPAGES pair', async () => {
    const parsed = await parseFooter(
      `<w:p>${simpleField("'", 'PAGE', '35')}${simpleField("'", 'NUMPAGES', '12')}</w:p>`,
    )
    expect(parsed.footerText).toBe(`${PAGE_MARK}${TOTAL_PAGES_MARK}`)
    expect(parsed.footerHasPageNumber).toBe(true)
  })

  it('keeps the cached result of a single-quoted non-page field', async () => {
    const parsed = await parseFooter(`<w:p>${complexField("'", 'DATE', '2026-01-01')}</w:p>`)
    expect(parsed.footerText).toBe('2026-01-01')
    expect(parsed.footerHasPageNumber).toBe(false)
  })

  it('keeps a single-quoted field inside surrounding text and its run styling', async () => {
    const parsed = await parseFooter(
      '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>' +
        `<w:r><w:rPr><w:rStyle w:val="PageNumber"/></w:rPr><w:fldChar w:fldCharType='begin'></w:fldChar></w:r>` +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        "<w:r><w:fldChar w:fldCharType='separate'></w:fldChar></w:r>" +
        '<w:r><w:t>35</w:t></w:r>' +
        "<w:r><w:fldChar w:fldCharType='end'></w:fldChar></w:r>" +
        '<w:r><w:t xml:space="preserve"> of </w:t></w:r>' +
        `${complexField("'", 'NUMPAGES', '12')}</w:p>`,
    )
    expect(parsed.footerText).toBe(`Page ${PAGE_MARK} of ${TOTAL_PAGES_MARK}`)
    expect(parsed.footerHasPageNumber).toBe(true)
  })

  it('reads a single-quoted PAGE field in a header and on save', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId91" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: hdrPart(`<w:p>${complexField("'", 'PAGE', '9')}</w:p>`),
          contentType: HDR_TYPE,
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId91"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.headerText).toBe(PAGE_MARK)
    expect(parsed.headerHasPageNumber).toBe(true)
  })
})

describe('cached header/footer field result run boundaries', () => {
  it('keeps canonical and simple cached fields alongside live page fields', async () => {
    const parsed = await parseFooter(
      `<w:p>${complexField('"', 'SECTIONPAGES', '17')}` +
        `${simpleField('"', 'SECTIONPAGES', '18')}` +
        `<w:r><w:fldChar w:fldCharType="begin"/><w:instrText> PAGE </w:instrText>` +
        `<w:fldChar w:fldCharType="separate"/><w:t>99</w:t><w:fldChar w:fldCharType="end"/></w:r>` +
        `<w:r><w:fldChar w:fldCharType="begin"/><w:instrText> NUMPAGES </w:instrText>` +
        `<w:fldChar w:fldCharType="separate"/><w:t>99</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    )
    expect(parsed.footerText).toBe(`1718${PAGE_MARK}${TOTAL_PAGES_MARK}`)
  })

  it('does not expose instruction-side text or invent a missing cache', async () => {
    const parsed = await parseFooter(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> DATE </w:instrText>` +
        `<w:t>hidden</w:t><w:fldChar w:fldCharType="separate"/><w:t>shown</w:t>` +
        `<w:fldChar w:fldCharType="end"/></w:r>` +
        `<w:r><w:fldChar w:fldCharType="begin"/><w:instrText> DATE </w:instrText>` +
        `<w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    )
    expect(parsed.footerText).toBe('shown')
  })

  it.each(['SECTIONPAGES', 'DATE', 'FILENAME'])(
    'keeps a same-run %s cache and its formatting without evaluating it',
    async (instr) => {
      const parsed = await parseFooter(
        `<w:p><w:r><w:t>Before </w:t></w:r>` +
          `<w:r><w:rPr><w:b/><w:color w:val="123456"/></w:rPr>` +
          `<w:fldChar w:fldCharType="begin"/><w:instrText> ${instr} </w:instrText>` +
          `<w:fldChar w:fldCharType="separate"/><w:t>17</w:t>` +
          `<w:fldChar w:fldCharType="end"/></w:r>` +
          `<w:r><w:t> after</w:t></w:r></w:p>`,
      )
      expect(parsed.footerText).toBe('Before 17 after')
      expect(parsed.footerParas![0].runs.find((run) => run.text === '17')).toMatchObject({
        bold: true,
        color: '123456',
      })
      const saved = await saveDocx(
        parsed,
        parsed.blocks
          .filter((block) => !block.hidden)
          .map((block) => ({ kind: 'original', docxIndex: block.docxIndex! })),
      )
      expect(saved).toBe(parsed.internal.originalBytes)
    },
  )

  it('keeps partial edge runs and complete middle result runs in order', async () => {
    const parsed = await parseFooter(
      `<w:p><w:r><w:rPr><w:b/></w:rPr>` +
        `<w:fldChar w:fldCharType="begin"/><w:instrText> SECTIONPAGES </w:instrText>` +
        `<w:fldChar w:fldCharType="separate"/><w:t>1</w:t></w:r>` +
        `<w:r><w:rPr><w:i/></w:rPr><w:t>2</w:t></w:r>` +
        `<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>3</w:t>` +
        `<w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    )
    expect(parsed.footerText).toBe('123')
    expect(parsed.footerParas![0].runs).toMatchObject([
      { text: '1', bold: true },
      { text: '2', italic: true },
      { text: '3', underline: true },
    ])
  })
})
