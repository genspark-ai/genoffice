import { describe, expect, it } from 'vitest'
import { outerRowsXml, tableRowFlags } from '../src/renderer/pagination'

const tbl = (rows: string[]) => `<w:tbl><w:tblPr/>${rows.join('')}</w:tbl>`
const tr = (trPr: string, cell: string) =>
  `<w:tr><w:trPr>${trPr}</w:trPr><w:tc>${cell}</w:tc></w:tr>`

describe('tableRowFlags: nested tables are cell content', () => {
  it('a nested table inside a cell does not shift the outer rows flags', () => {
    const nested = tbl([tr('<w:tblHeader/>', '<w:p/>'), tr('<w:cantSplit/>', '<w:p/>')])
    const xml = tbl([
      tr('', `<w:p/>${nested}`),
      tr('<w:trHeight w:val="440"/>', '<w:p/>'),
      tr('<w:cantSplit/>', '<w:p><w:pPr><w:keepNext/></w:pPr></w:p>'),
    ])
    expect(tableRowFlags(xml)).toEqual([
      { isHeader: false, cantSplit: false },
      { isHeader: false, cantSplit: false, minHPx: 440 / 15 },
      { isHeader: false, cantSplit: true, keepNext: true },
    ])
  })

  it('outerRowsXml keeps the outer markup and drops doubly nested tables too', () => {
    const inner = tbl([tr('', '<w:p/>')])
    const nested = tbl([tr('', `<w:p/>${inner}`)])
    const xml = tbl([tr('<w:tblHeader/>', `<w:p><w:r><w:t>a</w:t></w:r></w:p>${nested}`)])
    const outer = outerRowsXml(xml)
    expect(outer).toContain('<w:tblHeader/>')
    expect(outer).toContain('<w:t>a</w:t>')
    expect((outer.match(/<w:tr>/g) ?? []).length).toBe(1)
    expect((outer.match(/<w:tbl>/g) ?? []).length).toBe(1)
  })
})
