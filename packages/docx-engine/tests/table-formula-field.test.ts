import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
const sep = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const instr = (text: string) =>
  `<w:r><w:instrText xml:space="preserve"> ${text} </w:instrText></w:r>`
const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const cell = (inner: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p>${inner}</w:p></w:tc>`
const table = (...cells: string[]) =>
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  cells.map((inner) => `<w:tr>${cell(inner)}</w:tr>`).join('') +
  '</w:tbl><w:p/>'

describe('a formula field in a table cell', () => {
  const formula = begin + instr('=SUM(ABOVE) \\# "0.00"') + sep + t('10.00') + end

  it('is read as a field run whose text is the cached result', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: table(t('4'), t('6'), formula) }))
    const cells = doc.blocks[0].table!.rows.map((row) => row[0])
    expect(cells[2].richParas?.[0].runs).toEqual([
      expect.objectContaining({ text: '10.00', instrField: '=SUM(ABOVE) \\# "0.00"' }),
    ])
  })

  it('keeps a formula built from another field as the runs it was', async () => {
    const nested =
      begin +
      instr('=') +
      begin +
      instr('PAGE') +
      sep +
      t('2') +
      end +
      instr('+1') +
      sep +
      t('3') +
      end
    const doc = await parseDocx(await buildDocx({ bodyXml: table(t('4'), nested) }))
    const runs = doc.blocks[0].table!.rows[1][0].richParas?.[0].runs ?? []
    expect(runs.some((run) => run.instrField !== undefined)).toBe(false)
    expect(runs.map((run) => run.text).join('')).toBe('3')
  })
})
