import { describe, expect, it } from 'vitest'
import { generateTableModelXml, parseDocx, patchTableCellTexts } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const TABLE_XML =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr>' +
  '<w:tc><w:tcPr><w:shd w:val="clear" w:fill="1F3864"/></w:tcPr>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>标题A</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:shd w:val="clear" w:fill="1F3864"/></w:tcPr>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>标题B</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '<w:tr>' +
  '<w:tc><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>乙</w:t></w:r></w:p><w:p><w:r><w:t>第二段</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '</w:tbl>'

describe('patchTableCellTexts', () => {
  it('replaces a single cell text and keeps the rest byte-identical', () => {
    const out = patchTableCellTexts(TABLE_XML, [null, [['已修改'], null]])
    expect(out).toContain('<w:t xml:space="preserve">已修改</w:t>')
    expect(out).not.toContain('<w:t>甲</w:t>')
    // untouched cells keep exact original bytes
    expect(out).toContain('<w:t>标题A</w:t>')
    expect(out).toContain('<w:t>标题B</w:t>')
    expect(out).toContain('<w:t>乙</w:t>')
  })

  it('preserves tcPr, first-paragraph pPr and first-run rPr of the edited cell', () => {
    const out = patchTableCellTexts(TABLE_XML, [[['新标题'], null], null])
    const cell = /<w:tc>[\s\S]*?<\/w:tc>/.exec(out)![0]
    expect(cell).toContain('<w:shd w:val="clear" w:fill="1F3864"/>')
    expect(cell).toContain('<w:jc w:val="center"/>')
    expect(cell).toContain('<w:b/>')
    expect(cell).toContain('<w:color w:val="FFFFFF"/>')
    expect(cell).toContain('<w:t xml:space="preserve">新标题</w:t>')
  })

  it('writes multiple paragraphs per cell and escapes XML entities', () => {
    const out = patchTableCellTexts(TABLE_XML, [null, [null, ['A<B', '&C']]])
    expect(out).toContain('<w:t xml:space="preserve">A&lt;B</w:t>')
    expect(out).toContain('<w:t xml:space="preserve">&amp;C</w:t>')
  })

  it('round-trips through the parser with the new text', async () => {
    const patched = patchTableCellTexts(TABLE_XML, [null, [['改甲'], ['改乙一', '改乙二']]])
    const doc = await parseDocx(await buildDocx({ bodyXml: patched }))
    const table = doc.blocks[0].table!
    expect(table.rows[1][0].paras).toEqual(['改甲'])
    expect(table.rows[1][1].paras).toEqual(['改乙一', '改乙二'])
    // header styles survived
    expect(table.rows[0][0]).toMatchObject({ fill: '1F3864', bold: true, align: 'center' })
  })

  it('skips cells containing nested tables', () => {
    const nested =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>外层</w:t></w:r></w:p>' +
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>内层</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      '<w:p/></w:tc></w:tr></w:tbl>'
    const out = patchTableCellTexts(nested, [[['不应生效']]])
    expect(out).toBe(nested)
  })

  it('handles vertical merges: continue cells stay untouched', () => {
    const merged =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>合并首</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>' +
      '</w:tbl>'
    const out = patchTableCellTexts(merged, [[['新合并首']], null])
    expect(out).toContain('<w:t xml:space="preserve">新合并首</w:t>')
    expect(out).toContain('<w:vMerge w:val="restart"/>')
    expect(out).toContain('<w:vMerge/>')
  })
})

describe('generateTableModelXml', () => {
  it('round-trips structural rows, columns, styling, and merged cells', async () => {
    const xml = generateTableModelXml({
      colWidthsPct: [25, 35, 40],
      rows: [
        [
          {
            paras: ['Merged header'],
            colSpan: 2,
            vMerge: 'restart',
            fill: '1F3864',
            color: 'FFFFFF',
            bold: true,
            align: 'center',
          },
          { paras: ['C'] },
        ],
        [{ paras: [''], colSpan: 2, vMerge: 'continue', fill: '1F3864' }, { paras: ['D'] }],
        [{ paras: ['A'] }, { paras: ['B'] }, { paras: ['C'] }],
      ],
    })
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    const model = parsed.blocks[0].table!

    expect(model.rows).toHaveLength(3)
    expect(model.rows[0][0]).toMatchObject({
      paras: ['Merged header'],
      colSpan: 2,
      vMerge: 'restart',
      fill: '1F3864',
      color: 'FFFFFF',
      bold: true,
      align: 'center',
    })
    expect(model.rows[1][0]).toMatchObject({ colSpan: 2, vMerge: 'continue' })
    expect(model.rows[2].map((cell) => cell.paras[0])).toEqual(['A', 'B', 'C'])
    expect(model.colWidthsPct?.map(Math.round)).toEqual([25, 35, 40])
  })

  it('reuses the imported tblPr while regenerating rows and rich runs', async () => {
    const template =
      '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>' +
      '<w:tblCellMar><w:left w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const xml = generateTableModelXml(
      {
        rows: [
          [
            {
              paras: ['Rich'],
              richParas: [
                {
                  runs: [
                    {
                      text: 'Rich',
                      bold: true,
                      color: 'C00000',
                      font: 'Arial',
                      sizeHalfPoints: 30,
                    },
                  ],
                },
              ],
            },
            { paras: ['New'] },
          ],
        ],
      },
      template,
    )
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    const run = parsed.blocks[0].table?.rows[0][0].richParas?.[0].runs[0]

    expect(xml).toContain('<w:tblStyle w:val="TableGrid"/>')
    expect(xml).toContain('<w:tblCellMar><w:left w:w="120" w:type="dxa"/></w:tblCellMar>')
    expect(run).toMatchObject({
      text: 'Rich',
      bold: true,
      color: 'C00000',
      font: 'Arial',
      sizeHalfPoints: 30,
    })
  })
})

describe('tcPr/trPr fidelity and new attributes', () => {
  it('generateTableModelXml: unmodeled rawTcPr attributes are kept; vAlign/borders are written from the model', async () => {
    const { generateTableModelXml } = await import('../src/index')
    const model = {
      rows: [
        [
          {
            paras: ['a'],
            vAlign: 'center' as const,
            borders: {
              top: { style: 'single', szEighths: 8, color: 'FF0000' },
              bottom: { style: 'none' },
            },
            rawTcPr:
              '<w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:tcMar><w:left w:w="200" w:type="dxa"/></w:tcMar><w:textDirection w:val="btLr"/></w:tcPr>',
          },
          { paras: ['b'] },
        ],
      ],
      rowHeightsTwips: [600],
      rawTrPrs: ['<w:trPr><w:cantSplit/></w:trPr>'],
    }
    const xml = generateTableModelXml(model)
    // Unmodeled properties are kept
    expect(xml).toContain('<w:tcMar><w:left w:w="200" w:type="dxa"/></w:tcMar>')
    expect(xml).toContain('<w:textDirection w:val="btLr"/>')
    // vAlign comes after textDirection (schema order); borders sit in the shd position range
    expect(xml).toContain('<w:vAlign w:val="center"/>')
    expect(xml.indexOf('<w:tcBorders>')).toBeLessThan(xml.indexOf('<w:tcMar>'))
    expect(xml).toContain('<w:top w:val="single" w:sz="8" w:space="0" w:color="FF0000"/>')
    expect(xml).toContain('<w:bottom w:val="none"/>')
    // trPr: cantSplit kept + trHeight injected
    expect(xml).toContain(
      '<w:trPr><w:cantSplit/><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr>',
    )
  })

  it('parse read-back: vAlign/tcBorders/trHeight/rawTcPr/rawTrPr', async () => {
    const { parseDocx } = await import('../src/index')
    const { buildDocx } = await import('./helpers/build-docx')
    const tbl =
      '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:trHeight w:val="500"/><w:cantSplit/></w:trPr>' +
      '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vAlign w:val="bottom"/>' +
      '<w:tcBorders><w:top w:val="dashed" w:sz="12" w:color="00FF00"/></w:tcBorders>' +
      '<w:tcMar><w:left w:w="100" w:type="dxa"/></w:tcMar></w:tcPr>' +
      '<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: tbl }))
    const table = parsed.blocks.find((b) => b.type === 'table')!
    const model = table.table!
    const cell = model.rows[0][0]
    expect(cell.vAlign).toBe('bottom')
    expect(cell.borders?.top).toEqual({ style: 'dashed', szEighths: 12, color: '00FF00' })
    expect(cell.rawTcPr).toContain('<w:tcMar>')
    expect(model.rowHeightsTwips).toEqual([500])
    expect(model.rawTrPrs?.[0]).toContain('<w:cantSplit/>')
  })

  it('parse clamps EMU-polluted trHeight to the Word maximum (31680 twips)', async () => {
    const { parseDocx } = await import('../src/index')
    const { buildDocx } = await import('./helpers/build-docx')
    const tbl =
      '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:trHeight w:val="504000" w:hRule="atLeast"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:trPr><w:trHeight w:val="500"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: tbl }))
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    expect(model.rowHeightsTwips).toEqual([31680, 500])
  })
})

describe('tblStyleId table style reference', () => {
  it('replaces/inserts/removes w:tblStyle, adding a default tblLook when missing', async () => {
    const { generateTableModelXml } = await import('../src/index')
    const model = { rows: [[{ paras: ['x'] }]], tblStyleId: 'FancyTable' }
    const orig =
      '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    const xml = generateTableModelXml(model, orig)
    expect(xml).toContain('<w:tblPr><w:tblStyle w:val="FancyTable"/><w:tblW')
    expect(xml.match(/<w:tblStyle/g)).toHaveLength(1)
    expect(xml).toContain('<w:tblLook w:val="04A0"')
    // Removal
    const removed = generateTableModelXml({ rows: [[{ paras: ['x'] }]], tblStyleId: '' }, orig)
    expect(removed).not.toContain('<w:tblStyle')
    // undefined leaves it untouched
    const kept = generateTableModelXml({ rows: [[{ paras: ['x'] }]] }, orig)
    expect(kept).toContain('<w:tblStyle w:val="TableGrid"/>')
  })
})
