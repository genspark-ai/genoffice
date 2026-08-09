import { describe, expect, it } from 'vitest'
import {
  generateTableModelXml,
  parseDocx,
  patchTableCellTexts,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { buildDocx, NESTED_TABLE_XML } from './helpers/build-docx'

const INNER = (label: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:p><w:r><w:t>Inner${label}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`

/** heading - table - heading - table - trailing empty p, all inside one cell */
const INTERLEAVED_TABLE_XML =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc>' +
  '<w:p><w:r><w:t>Title1</w:t></w:r></w:p>' +
  INNER('A') +
  '<w:p><w:r><w:t>Title2</w:t></w:r></w:p>' +
  INNER('B') +
  '<w:p/>' +
  '</w:tc></w:tr></w:tbl>'

describe('nested table editing (surgical text patch + regeneration)', () => {
  it('patches nested cell texts and keeps the outer cell bytes', () => {
    const out = patchTableCellTexts(NESTED_TABLE_XML, [
      [{ nested: [[[['EditedInnerA'], null]]] }, null],
    ])
    expect(out).toContain('EditedInnerA')
    expect(out).toContain('<w:t>InnerB</w:t>')
    expect(out).toContain('<w:t>Outer</w:t>')
    expect(out).toContain('<w:t>RightCell</w:t>')
    // outer structure is not rearranged
    expect(out.indexOf('<w:tbl>')).toBe(0)
  })

  it('plain-array cell patches still skip cells containing nested tables', () => {
    const out = patchTableCellTexts(NESTED_TABLE_XML, [[['ShouldNotApply'], null]])
    expect(out).toBe(NESTED_TABLE_XML)
  })

  it('regenerating a table model keeps nested tables (structure edits no longer drop them)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: NESTED_TABLE_XML }))
    const model = doc.blocks[0].table!
    const xml = generateTableModelXml(model, doc.blocks[0].originalXml ?? undefined)
    expect(xml).toContain('InnerA')
    expect(xml).toContain('InnerB')
    // nested table is followed by an empty paragraph (OOXML: tc must end with w:p)
    expect(/<\/w:tbl><w:p\/><\/w:tc>/.test(xml)).toBe(true)
    // reparsing still yields the nested structure
    const saved = await saveDocx(doc, [{ kind: 'xml', xml }])
    const reparsed = await parseDocx(saved)
    const nested = reparsed.blocks[0].table!.rows[0][0].nestedTables
    expect(nested?.[0]?.rows[0].map((c) => c.paras[0])).toEqual(['InnerA', 'InnerB'])
  })

  it('parses paragraph anchors for interleaved nested tables and regenerates in reading order', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: INTERLEAVED_TABLE_XML }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.paras).toEqual(['Title1', 'Title2'])
    expect(cell.nestedTableAnchors).toEqual([1, 2])

    const xml = generateTableModelXml(doc.blocks[0].table!)
    const order = ['Title1', 'InnerA', 'Title2', 'InnerB'].map((t) => xml.indexOf(t))
    expect(order.every((idx) => idx >= 0)).toBe(true)
    expect([...order].sort((a, b) => a - b)).toEqual(order)
    // trailing nested table is still followed by an empty paragraph
    expect(/<\/w:tbl><w:p\/><\/w:tc>/.test(xml)).toBe(true)

    // reparse of the regenerated table keeps the anchors
    const reparsed = await parseDocx(await saveDocx(doc, [{ kind: 'xml', xml }]))
    expect(reparsed.blocks[0].table!.rows[0][0].nestedTableAnchors).toEqual([1, 2])
  })

  it('untouched interleaved nested-table docx saves byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: INTERLEAVED_TABLE_XML })
    const doc = await parseDocx(bytes)
    const blocks: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toBe(bytes)
  })
})
