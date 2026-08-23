import { describe, expect, it } from 'vitest'
import { parseDocx, patchFieldParagraphXml, patchMathTokens } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const TOC_ENTRY =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:b/></w:rPr><w:t>First </w:t></w:r>' +
  '<w:r><w:t>chapter</w:t></w:r><w:r><w:tab/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>'

const FORMULA =
  '<w:p><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>' +
  '<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></w:p>'

describe('protected visible-text patching', () => {
  it('edits TOC title and page while preserving field structure and run formatting', async () => {
    const patched = patchFieldParagraphXml(TOC_ENTRY, {
      left: 'Updated & chapter',
      right: '12',
    })

    expect(patched).toContain('<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText>')
    expect(patched).toContain('<w:instrText> PAGEREF _Toc1 \\h </w:instrText>')
    expect(patched).toContain('<w:rPr><w:b/></w:rPr>')
    expect(patched).toContain('&amp; chapter')

    const parsed = await parseDocx(await buildDocx({ bodyXml: patched }))
    expect(parsed.blocks[0].fieldDisplay?.left).toBe('Updated & chapter')
    expect(parsed.blocks[0].fieldDisplay?.right).toBe('12')
  })

  it('patches the page number after the LAST tab, leaving the title and num intact', async () => {
    const entry =
      '<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr>' +
      '<w:r><w:t>1.1.</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>Latar Belakang</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>7</w:t></w:r></w:p>'
    const patched = patchFieldParagraphXml(entry, { left: 'Judul Baru', right: '9' })
    expect(patched).toContain('<w:t>1.1.</w:t>')
    expect(patched).toContain('<w:t>Judul Baru</w:t>')
    expect(patched).toContain('<w:t>9</w:t>')
    expect(patched).not.toContain('Latar Belakang')
  })

  it('edits OMML tokens while preserving the formula tree', async () => {
    const patched = patchMathTokens(FORMULA, ['x & 1', 'y'])
    expect(patched).toContain('<m:f><m:num>')
    expect(patched).toContain('<m:t>x &amp; 1</m:t>')
    expect(patched).toContain('<m:den><m:r><m:t>y</m:t>')

    const parsed = await parseDocx(await buildDocx({ bodyXml: patched }))
    expect(parsed.blocks[0].formulaDisplay?.tokens).toEqual(['x & 1', 'y'])
    expect(parsed.blocks[0].formulaDisplay?.mathml).toContain('<mfrac>')
    expect(parsed.blocks[0].previewText).toBe('x & 1y')
  })

  it('refuses a formula patch when token count changes', () => {
    expect(patchMathTokens(FORMULA, ['only-one'])).toBe(FORMULA)
  })
})
