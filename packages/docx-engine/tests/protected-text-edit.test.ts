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

  it('pins xml:space preserve when the patched text gains edge spaces', async () => {
    const entry =
      '<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr>' +
      '<w:r><w:t>Title</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>7</w:t></w:r></w:p>'
    const patched = patchFieldParagraphXml(entry, { left: ' Title ', right: '7' })
    // without preserve, wtText drops the edge spaces on re-parse and Word loses them too
    expect(patched).toContain('<w:t xml:space="preserve"> Title </w:t>')
    const parsed = await parseDocx(await buildDocx({ bodyXml: patched }))
    expect(parsed.blocks[0].fieldDisplay?.kind).toBe('tocLine')
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

  it('handles numeric char refs so distribution stays aligned', async () => {
    const entry =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:r><w:t>Foo &#8211; Bar</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>12</w:t></w:r></w:p>'
    const patched = patchFieldParagraphXml(entry, { left: 'Foo – Baz', right: '13' })
    const parsed = await parseDocx(await buildDocx({ bodyXml: patched }))
    expect(parsed.blocks[0].fieldDisplay?.left).toBe('Foo – Baz')
    expect(parsed.blocks[0].fieldDisplay?.right).toBe('13')
  })

  it('pins preserve for NBSP and other whitespace at edges (not just tab/space)', () => {
    const entry =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:r><w:t>Title</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>7</w:t></w:r></w:p>'
    const nbsp = '\u00A0Title\u00A0'
    const patched = patchFieldParagraphXml(entry, { left: nbsp, right: '7' })
    expect(patched).toContain('xml:space="preserve"')
    expect(patched).toContain(`<w:t xml:space="preserve">${nbsp}</w:t>`)
    expect(patched).toContain('\u00A0Title\u00A0')
  })

  it('patches a self-closing empty run (<w:t/>) instead of no-oping', async () => {
    const entry =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:r><w:t/></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>7</w:t></w:r></w:p>'
    const patched = patchFieldParagraphXml(entry, { left: 'Hello', right: '7' })
    expect(patched).toContain('<w:t>Hello</w:t>')
    const parsed = await parseDocx(await buildDocx({ bodyXml: patched }))
    expect(parsed.blocks[0].fieldDisplay?.left).toBe('Hello')
  })

  it('pins preserve when a self-closing run gains edge whitespace', () => {
    const entry =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:r><w:t/></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>7</w:t></w:r></w:p>'
    const patched = patchFieldParagraphXml(entry, { left: '\u00A0Hello', right: '7' })
    expect(patched).toContain('<w:t xml:space="preserve">\u00A0Hello</w:t>')
  })

  it('ignores self-closing <m:t/> so the token count matches mathTokensOf', () => {
    const withEmpty = FORMULA.replace('<m:den>', '<m:den><m:r><m:t/></m:r>')
    const patched = patchMathTokens(withEmpty, ['a', 'b'])
    expect(patched).toContain('<m:t>a</m:t>')
    expect(patched).toContain('<m:t>b</m:t>')
    expect(patched).toContain('<m:t/>')
  })

  it('leaves out-of-range and surrogate char refs untouched instead of throwing', () => {
    const base =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:r><w:t>REFTEXT</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:t>12</w:t></w:r></w:p>'
    for (const ref of ['&#99999999;', '&#xFFFFFFFF;', '&#55296;', '&#xD800;']) {
      const entry = base.replace('REFTEXT', `A ${ref} B`)
      expect(() => patchFieldParagraphXml(entry, { left: 'A X B', right: '13' })).not.toThrow()
      const patched = patchFieldParagraphXml(entry, { left: 'A X B', right: '13' })
      expect(patched).toContain('A X B')
    }
  })
})
