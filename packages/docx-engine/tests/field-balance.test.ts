import { describe, expect, it } from 'vitest'
import { balanceFieldChars } from '../src/field-balance'

const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
const instr = '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL </w:instrText></w:r>'
const sep = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const t = (s: string) => `<w:r><w:t>${s}</w:t></w:r>`
const p = (inner: string) => `<w:p>${inner}</w:p>`

const count = (xml: string, type: string) =>
  (xml.match(new RegExp(`w:fldCharType="${type}"`, 'g')) ?? []).length

describe('balanceFieldChars', () => {
  it('returns balanced bodies untouched, including fields spanning paragraphs', () => {
    const body = p(begin + instr + sep + t('Alpha')) + p(t('Beta')) + p(t('Gamma') + end)
    expect(balanceFieldChars(body)).toBe(body)
    const nested = p(begin + instr + sep + begin + instr + sep + t('x') + end + t('y') + end)
    expect(balanceFieldChars(nested)).toBe(nested)
    expect(balanceFieldChars(p(t('plain')))).toBe(p(t('plain')))
  })

  it('closes a begin whose end paragraph was deleted at the end of its own paragraph', () => {
    const body = p(begin + instr + sep + t('Alpha')) + p(t('Beta'))
    const out = balanceFieldChars(body)
    expect(out).toBe(p(begin + instr + sep + t('Alpha') + end) + p(t('Beta')))
  })

  it('drops end and separate runs that have no open field', () => {
    const body = p(t('Beta')) + p(t('Gamma') + end) + p(sep + t('x'))
    expect(balanceFieldChars(body)).toBe(p(t('Beta')) + p(t('Gamma')) + p(t('x')))
  })

  it('keeps a stray end run with attributes and rPr out as a whole run', () => {
    const body = p(
      '<w:r w:rsidR="00AB"><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>' + t('x'),
    )
    expect(balanceFieldChars(body)).toBe(p(t('x')))
  })

  it('pairs a duplicated begin from a split paragraph with the real end and closes the first', () => {
    const body =
      p(begin + instr + sep + t('Alpha')) + p(begin + instr + sep + t('Beta')) + p(t('Gamma') + end)
    const out = balanceFieldChars(body)
    expect(count(out, 'begin')).toBe(2)
    expect(count(out, 'end')).toBe(2)
    expect(out.startsWith(p(begin + instr + sep + t('Alpha') + end))).toBe(true)
  })

  it('closes at the outer paragraph, not inside a nested textbox paragraph', () => {
    const inner = '<w:txbxContent>' + p(t('box')) + '</w:txbxContent>'
    const body = p(begin + instr + sep + t('Alpha') + '<w:r><w:pict>' + inner + '</w:pict></w:r>')
    const out = balanceFieldChars(body)
    expect(out.endsWith(end + '</w:p>')).toBe(true)
    expect(out).toContain(p(t('box')))
  })

  it('treats a self-closing empty paragraph as a closed paragraph', () => {
    const body = '<w:p/>' + p(begin + instr + sep + t('Alpha'))
    expect(balanceFieldChars(body)).toBe('<w:p/>' + p(begin + instr + sep + t('Alpha') + end))
  })

  it('balances single-quoted fldCharType runs', () => {
    const qbegin = "<w:r><w:fldChar w:fldCharType='begin'/></w:r>"
    const qen = "<w:r><w:fldChar w:fldCharType='end'/></w:r>"
    const body = p(qbegin + instr + sep + t('Alpha')) + p(t('Beta'))
    const out = balanceFieldChars(body)
    expect(out).toContain('Alpha')
    expect(out).toContain('w:fldCharType="end"')
    expect(balanceFieldChars(p(t('x')) + p(qen))).toBe(p(t('x')) + p(''))
  })
})
