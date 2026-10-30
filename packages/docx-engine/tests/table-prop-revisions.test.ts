import { describe, expect, it } from 'vitest'
import {
  generateTableModelXml,
  parseDocx,
  resolvePropChangeXml,
  saveDocx,
  type TableModel,
} from '../src/index'
import { buildDocx, PROP_REVISION_TABLE_XML, PROP_REVISION_AT as AT } from './helpers/build-docx'

const SECT_XML =
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange w:id="31" ${AT} w:original="1."/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:sectPrChange w:id="32" ${AT}><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p>` +
  '<w:p><w:r><w:t>Tail</w:t></w:r></w:p>'

describe('table / numbering / section property revisions', () => {
  it('parses every property-change record into the model', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PROP_REVISION_TABLE_XML }))
    const model = doc.blocks[0].table!
    expect(model.tblPrChange).toEqual({
      author: 'Alice',
      date: '2026-01-01T00:00:00Z',
      id: '21',
      oldXml: '<w:tblW w:w="0" w:type="auto"/><w:jc w:val="left"/>',
    })
    expect(model.tblGridChange).toMatchObject({
      id: '22',
      oldXml: '<w:gridCol w:w="3000"/><w:gridCol w:w="5000"/>',
    })
    expect(model.rowTblPrExChanges?.[0]).toMatchObject({ id: '23' })
    expect(model.rawTblPrExs?.[0]).toMatch(/^<w:tblPrEx>.*<\/w:tblPrEx>$/)
    expect(model.rowPrChanges).toEqual([
      {
        author: 'Alice',
        date: '2026-01-01T00:00:00Z',
        id: '24',
        oldXml: '<w:trHeight w:val="300"/>',
      },
      null,
      null,
    ])
    expect(model.rows[0][0].tcPrChange).toMatchObject({
      id: '25',
      oldXml: '<w:tcW w:w="4000" w:type="dxa"/>',
    })
    expect(model.rows[2][0].cellMerge).toEqual({
      author: 'Alice',
      date: '2026-01-01T00:00:00Z',
      id: '26',
      vMerge: 'cont',
    })
  })

  it('keeps the untouched table byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: PROP_REVISION_TABLE_XML })
    const doc = await parseDocx(bytes)
    expect(await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }])).toEqual(bytes)
  })

  it('a regenerated table with pending records keeps every record', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PROP_REVISION_TABLE_XML }))
    const xml = generateTableModelXml(doc.blocks[0].table!, doc.blocks[0].originalXml ?? undefined)
    for (const name of [
      'tblPrChange',
      'tblGridChange',
      'tblPrExChange',
      'trPrChange',
      'tcPrChange',
      'cellMerge',
    ])
      expect(xml.match(new RegExp(`<w:${name}\\b`, 'g')), name).toHaveLength(1)
    expect(xml).toContain('<w:tr><w:tblPrEx>')
  })

  it('a continue cell rebuilt without its bytes re-emits a pending w:cellMerge', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PROP_REVISION_TABLE_XML }))
    const src = doc.blocks[0].table!
    expect(src.rows[2][0].cellMerge).toMatchObject({ vMerge: 'cont' })
    // the editor rebuilds merged-away cells from the restart cell: same record, restart bytes
    const restart = src.rows[1][0]
    const rows = src.rows.map((row, ri) =>
      ri === 2
        ? [
            {
              ...restart,
              paras: [''],
              vMerge: 'continue' as const,
              cellMerge: src.rows[2][0].cellMerge,
            },
            row[1],
          ]
        : row,
    )
    const xml = generateTableModelXml({ ...src, rows }, doc.blocks[0].originalXml ?? undefined)
    expect(xml.match(/<w:cellMerge\b/g)).toHaveLength(1)
    expect(xml).toContain(
      '<w:vMerge/><w:cellMerge w:id="26" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:vMerge="cont"/>',
    )
    const accepted = generateTableModelXml(
      {
        ...src,
        rows: rows.map((row) => row.map((c) => (c.cellMerge ? { ...c, cellMerge: null } : c))),
      },
      doc.blocks[0].originalXml ?? undefined,
    )
    expect(accepted).not.toContain('cellMerge')
  })

  it('accept drops the records and keeps the current properties', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PROP_REVISION_TABLE_XML }))
    const model: TableModel = {
      ...doc.blocks[0].table!,
      tblPrChange: null,
      tblGridChange: null,
      rowPrChanges: [null, null, null],
      rowTblPrExChanges: [null, null, null],
    }
    model.rows = model.rows.map((row) =>
      row.map((c) => ({
        ...c,
        ...(c.tcPrChange ? { tcPrChange: null } : {}),
        ...(c.cellMerge ? { cellMerge: null } : {}),
      })),
    )
    const xml = generateTableModelXml(model, doc.blocks[0].originalXml ?? undefined)
    expect(xml).not.toMatch(/Change\b|cellMerge/)
    expect(xml).toContain('<w:jc w:val="center"/>')
    expect(xml).toContain('<w:trHeight w:val="600"')
    expect(xml).toContain('w:fill="FF0000"')
    expect(xml).toContain('<w:tblCellMar><w:left w:w="200"')
    const reparsed = await parseDocx(await saveDocx(doc, [{ kind: 'xml', xml }]))
    const table = reparsed.blocks[0].table!
    expect(table.tblPrChange).toBeUndefined()
    expect(table.rowPrChanges).toBeUndefined()
    expect(table.rows[0][0].tcPrChange).toBeUndefined()
    expect(table.rows[2][0].cellMerge).toBeUndefined()
  })

  it('reject restores the recorded previous properties', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PROP_REVISION_TABLE_XML }))
    const src = doc.blocks[0].table!
    const model: TableModel = {
      ...src,
      align: undefined,
      // the editor restores the modeled subset of the old properties itself
      rowHeightsTwips: [300, null, null],
      rowHeightRules: ['atLeast', null, null],
      tblPrChange: { ...src.tblPrChange!, resolved: 'reject' },
      rowPrChanges: [{ ...src.rowPrChanges![0]!, resolved: 'reject' }, null, null],
      rowTblPrExChanges: [{ ...src.rowTblPrExChanges![0]!, resolved: 'reject' }, null, null],
    }
    model.rows = model.rows.map((row) =>
      row.map((c) => ({
        ...c,
        fill: undefined,
        ...(c.tcPrChange ? { tcPrChange: { ...c.tcPrChange, resolved: 'reject' as const } } : {}),
      })),
    )
    const xml = generateTableModelXml(model, doc.blocks[0].originalXml ?? undefined)
    expect(xml).not.toMatch(/PrChange|PrExChange/)
    expect(xml).toContain('<w:jc w:val="left"/>')
    expect(xml).not.toContain('<w:jc w:val="center"/>')
    expect(xml).toContain('<w:trHeight w:val="300"')
    expect(xml).not.toContain('w:val="600"')
    expect(xml).not.toContain('w:fill="FF0000"')
    // the exception's previous state was empty: the tblPrEx goes
    expect(xml).not.toContain('<w:tblPrEx')
  })

  it('numberingChange and sectPrChange are parsed on paragraphs and the trailing sectPr', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: SECT_XML, withNumbering: true }))
    expect(doc.blocks[0].numberingChange).toEqual({
      author: 'Alice',
      date: '2026-01-01T00:00:00Z',
      id: '31',
    })
    expect(doc.blocks[1].sectPrChange).toMatchObject({ id: '32' })
    const trailing = doc.blocks.find((b) => b.hidden && b.originalXml?.includes('<w:sectPr'))
    expect(trailing?.sectPrChange).toBeUndefined()
  })

  it('resolvePropChangeXml: numberingChange drops either way, sectPrChange reverts', () => {
    const pPr = /<w:pPr>.*?<\/w:pPr>/.exec(SECT_XML)![0]
    expect(resolvePropChangeXml(pPr, 'w:numberingChange', 'accept')).not.toContain(
      'numberingChange',
    )
    expect(resolvePropChangeXml(pPr, 'w:numberingChange', 'reject')).toContain(
      '<w:numId w:val="1"/></w:numPr>',
    )
    const sect = /<w:pPr><w:sectPr>.*?<\/w:pPr>/.exec(SECT_XML)![0]
    expect(resolvePropChangeXml(sect, 'w:sectPrChange', 'accept')).toBe(
      '<w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>',
    )
    expect(resolvePropChangeXml(sect, 'w:sectPrChange', 'reject')).toBe(
      '<w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>',
    )
  })

  it('the trailing sectPr record resolves through the save option', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        sectPrExtra: `<w:sectPrChange w:id="40" ${AT}><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange>`,
      }),
    )
    const trailing = doc.blocks.find((b) => b.hidden)!
    expect(trailing.sectPrChange).toMatchObject({ id: '40', author: 'Alice' })
    const xmlOf = async (mode: 'accept' | 'reject') => {
      const bytes = await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }], {
        trailingSectPrChange: mode,
      })
      const reparsed = await parseDocx(bytes)
      return reparsed.blocks.find((b) => b.hidden)!.originalXml!
    }
    const accepted = await xmlOf('accept')
    expect(accepted).not.toContain('sectPrChange')
    expect(accepted).toContain('w:w="11906"')
    const rejected = await xmlOf('reject')
    expect(rejected).not.toContain('sectPrChange')
    expect(rejected).toContain('w:w="12240"')
    expect(rejected).not.toContain('w:pgMar')
  })
})
