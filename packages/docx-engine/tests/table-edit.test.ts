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

  it('writes w:jc after w:tblW for an explicit alignment and round-trips it', async () => {
    const xml = generateTableModelXml({
      align: 'center',
      rows: [[{ paras: ['a'] }, { paras: ['b'] }]],
    })
    expect(xml).toMatch(/<w:tblW[^>]*\/><w:jc w:val="center"\/>/)
    const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(parsed.blocks[0].table!.align).toBe('center')
  })

  it("align 'left' strips the original w:jc; undefined keeps it", () => {
    const template =
      '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:jc w:val="right"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const model = { rows: [[{ paras: ['New'] }]] }
    expect(generateTableModelXml({ ...model, align: 'left' as const }, template)).not.toContain(
      '<w:jc',
    )
    expect(generateTableModelXml(model, template)).toContain('<w:jc w:val="right"/>')
    expect(generateTableModelXml({ ...model, align: 'center' as const }, template)).toContain(
      '<w:jc w:val="center"/>',
    )
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

describe('cell paragraph properties survive a table rebuild', () => {
  /** Word writes an RTL cell as w:bidi plus a LOGICAL w:jc: "left" means start, i.e. flush right. */
  const RTL_TABLE =
    '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:bidiVisual/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr>' +
    '<w:tc><w:p><w:pPr><w:bidi/><w:jc w:val="left"/><w:spacing w:after="80"/>' +
    '<w:ind w:left="120"/><w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/>' +
    '</w:pPr><w:r><w:t>الاسم</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>القيمة</w:t></w:r></w:p></w:tc>' +
    '</w:tr></w:tbl>'

  const rebuild = async (tableXml: string) => {
    const doc = await parseDocx(await buildDocx({ bodyXml: tableXml }))
    const model = doc.blocks.find((b) => b.table)?.table
    expect(model).toBeDefined()
    return generateTableModelXml(model!)
  }

  it('keeps w:bidi and writes the logical w:jc back', async () => {
    const out = await rebuild(RTL_TABLE)
    expect(out.match(/<w:bidi\/>/g)).toHaveLength(2)
    // parsed as visual "right", so Word's logical value on the way out is "left"
    expect(out.match(/<w:jc w:val="[^"]*"\/>/g)).toEqual(['<w:jc w:val="left"/>'])
  })

  it('leaves a bidi cell that declared no alignment without one', async () => {
    const out = await rebuild(RTL_TABLE)
    // the second cell had w:bidi and no w:jc; dropping w:bidi would flip it from the RTL
    // default (flush right) to the LTR default (flush left)
    const second = out.slice(out.lastIndexOf('<w:tc>'))
    expect(second).toContain('<w:bidi/>')
    expect(second).not.toContain('<w:jc ')
  })

  it('keeps spacing, indent and shading', async () => {
    const out = await rebuild(RTL_TABLE)
    expect(out).toContain('<w:spacing w:after="80"/>')
    expect(out).toContain('<w:ind w:left="120"/>')
    expect(out).toContain('<w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/>')
  })

  it("does not stamp another paragraph's alignment onto one that declared none", async () => {
    // mixed jc: no cell-level alignment, and the jc-less paragraph must stay that way
    const out = await rebuild(
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>b</w:t></w:r></w:p>' +
        '</w:tc></w:tr></w:tbl>',
    )
    expect(out.match(/<w:jc w:val="center"\/>/g)).toHaveLength(1)
    const second = out.slice(out.lastIndexOf('<w:p>'))
    expect(second).not.toContain('<w:jc ')
  })

  it('keeps a bidi paragraph free of alignment it never declared', async () => {
    const out = await rebuild(
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>عربي</w:t></w:r></w:p>' +
        '</w:tc></w:tr></w:tbl>',
    )
    const second = out.slice(out.lastIndexOf('<w:p>'))
    expect(second).toContain('<w:bidi/>')
    expect(second).not.toContain('<w:jc ')
  })

  it("writes every paragraph's jc back when they all agree (cell.align set)", async () => {
    const out = await rebuild(
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>' +
        '</w:tc></w:tr></w:tbl>',
    )
    expect(out.match(/<w:jc w:val="center"\/>/g)).toHaveLength(2)
  })
})

describe('trailing empty cell paragraph size survives regeneration', () => {
  it('writes the empty paragraph and its w:sz back', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>' +
          '<w:p><w:r><w:t>内容</w:t></w:r></w:p>' +
          '<w:p><w:pPr><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>' +
          '</w:tc></w:tr></w:tbl>',
      }),
    )
    const out = generateTableModelXml(doc.blocks[0].table!)
    expect(out.match(/<w:p[\s>]|<w:p\/>/g)).toHaveLength(2)
    expect(out).toContain('<w:pPr><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr>')
  })
})
