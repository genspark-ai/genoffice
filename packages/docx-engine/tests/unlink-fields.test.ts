import { describe, expect, it } from 'vitest'
import { docxWithBody, parseDocx, unlinkFieldsXml } from '../src'
import { buildDocx } from './helpers/build-docx'

const field = (instr: string, result: string, lock = false) =>
  `<w:r><w:fldChar w:fldCharType="begin"${lock ? ' w:fldLock="true"' : ''}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  result +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

describe('unlinking fields kept as XML', () => {
  it('leaves the result runs and drops the field and the hyperlink around it', () => {
    const line =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:hyperlink w:anchor="_Toc1" w:history="1"><w:r><w:t>Alpha</w:t></w:r><w:r><w:tab/></w:r>' +
      field('PAGEREF _Toc1 \\h', '<w:r><w:rPr><w:noProof/></w:rPr><w:t>3</w:t></w:r>') +
      '</w:hyperlink></w:p>'
    expect(unlinkFieldsXml(line)).toBe(
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
        '<w:r><w:t>Alpha</w:t></w:r><w:r><w:tab/></w:r>' +
        '<w:r><w:rPr><w:noProof/></w:rPr><w:t>3</w:t></w:r></w:p>',
    )
    const last = '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    expect(unlinkFieldsXml(last)).toBe('<w:p></w:p>')
    const plain = '<w:p><w:r><w:t>Body</w:t></w:r></w:p>'
    expect(unlinkFieldsXml(plain)).toBe(plain)
  })

  it('a package with that body reads as ordinary paragraphs', async () => {
    const source = await buildDocx({
      bodyXml: `<w:p>${field('TOC \\o', '<w:r><w:t>Alpha</w:t></w:r>')}</w:p><w:p><w:r><w:t>Body</w:t></w:r></w:p>`,
    })
    const bytes = await docxWithBody(
      source,
      unlinkFieldsXml(`<w:p>${field('TOC \\o', '<w:r><w:t>Alpha</w:t></w:r>')}</w:p>`),
    )
    const parsed = await parseDocx(bytes!)
    expect(
      parsed.blocks
        .filter((b) => !b.hidden)
        .map((b) => [b.type, b.runs?.map((r) => r.text).join('')]),
    ).toEqual([['paragraph', 'Alpha']])
  })
})
