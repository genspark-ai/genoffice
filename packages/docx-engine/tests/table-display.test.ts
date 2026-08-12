import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLED_TABLE_XML =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  // header row: shaded cells, bold white text, centered
  '<w:tr>' +
  '<w:tc><w:tcPr><w:shd w:val="clear" w:fill="1F3864"/></w:tcPr>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>排名</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:shd w:val="clear" w:fill="1F3864"/></w:tcPr>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>品牌</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:shd w:val="clear" w:fill="1F3864"/></w:tcPr>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>份额</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  // data row with colSpan
  '<w:tr>' +
  '<w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>比亚迪</w:t></w:r></w:p><w:p><w:r><w:t>第二段</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  // vertical merge pair
  '<w:tr>' +
  '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>竖排</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '<w:tr>' +
  '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>' +
  '<w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>d</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '</w:tbl>'

describe('table display model', () => {
  it('extracts rows, shading, bold, alignment, spans and column widths', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: STYLED_TABLE_XML }))
    const block = doc.blocks[0]
    expect(block.type).toBe('table')
    const table = block.table!
    expect(table.rows).toHaveLength(4)

    const header = table.rows[0]
    expect(header.map((c) => c.paras[0])).toEqual(['排名', '品牌', '份额'])
    expect(header[0]).toMatchObject({
      fill: '1F3864',
      color: 'FFFFFF',
      bold: true,
      align: 'center',
    })

    const dataRow = table.rows[1]
    expect(dataRow[1].colSpan).toBe(2)
    expect(dataRow[1].paras).toEqual(['比亚迪', '第二段'])
    expect(dataRow[0].bold).toBeUndefined()

    expect(table.rows[2][0].vMerge).toBe('restart')
    expect(table.rows[3][0].vMerge).toBe('continue')

    expect(table.colWidthsPct).toHaveLength(3)
    expect(Math.round(table.colWidthsPct![0])).toBe(25)
    expect(Math.round(table.colWidthsPct![1])).toBe(50)
  })

  it('original table XML is still saved byte-identical (display model is read-only)', async () => {
    const bytes = await buildDocx({ bodyXml: STYLED_TABLE_XML })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].originalXml).toContain('<w:tbl>')
  })

  it('parses table-level borders / alignment / cell margins / nested tables', async () => {
    const xml =
      '<w:tbl><w:tblPr>' +
      '<w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/><w:tblInd w:w="120" w:type="dxa"/>' +
      '<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="FF0000"/>' +
      '<w:insideH w:val="none"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>' +
      '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/></w:tblCellMar>' +
      '</w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>外层</w:t></w:r></w:p>' +
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>内A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>内B</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl><w:p/></w:tc>' +
      '<w:tc><w:p><w:r><w:t>右格</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const table = doc.blocks[0].table!
    expect(table.align).toBe('center')
    // tblInd is still parsed with jc=center (the renderer only applies it for left alignment)
    expect(table.indentTwips).toBe(120)
    expect(table.borders?.top).toMatchObject({ style: 'single', szEighths: 12, color: 'FF0000' })
    expect(table.borders?.insideH).toMatchObject({ style: 'none' })
    expect(table.borders?.insideV).toMatchObject({ style: 'single', szEighths: 4 })
    expect(table.cellMarTwips).toEqual({ top: 0, left: 0 })
    expect(table.colWidthsTwips).toEqual([3000, 3000])
    const outerCell = table.rows[0][0]
    expect(outerCell.paras[0]).toBe('外层')
    const nested = outerCell.nestedTables![0]
    expect(nested.rows[0].map((c) => c.paras[0])).toEqual(['内A', '内B'])
    expect(nested.colWidthsTwips).toEqual([1000, 1000])
  })

  it('parses cell textDirection and falls back to tcW when tblGrid is unusable', async () => {
    const xml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="0"/><w:gridCol w:w="0"/></w:tblGrid>' +
      '<w:tr>' +
      '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:textDirection w:val="tbRlV"/></w:tcPr><w:p><w:r><w:t>竖</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:textDirection w:val="btLr"/></w:tcPr><w:p><w:r><w:t>横转</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const table = doc.blocks[0].table!
    expect(table.colWidthsTwips).toEqual([2000, 6000])
    expect(Math.round(table.colWidthsPct![0])).toBe(25)
    expect(table.rows[0][0].textDirection).toBe('tbRl')
    expect(table.rows[0][1].textDirection).toBe('btLr')
  })

  it('merges legacy w:hMerge continue cells into the restart cell', async () => {
    const xml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr>' +
      '<w:tc><w:tcPr><w:hMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>合并头</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:hMerge/></w:tcPr><w:p/></w:tc>' +
      '<w:tc><w:p><w:r><w:t>右</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const row = doc.blocks[0].table!.rows[0]
    expect(row).toHaveLength(2)
    expect(row[0].colSpan).toBe(2)
    expect(row[0].paras[0]).toBe('合并头')
    expect(row[1].paras[0]).toBe('右')
  })

  it('inherits tblBorders / tblCellMar from the referenced table style', async () => {
    const styleXml =
      '<w:style w:type="table" w:styleId="GridStyled"><w:name w:val="Grid Styled"/>' +
      '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="16" w:color="00FF00"/>' +
      '<w:insideH w:val="dashed" w:sz="4"/></w:tblBorders>' +
      '<w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    const xml =
      '<w:tbl><w:tblPr><w:tblStyle w:val="GridStyled"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml, extraStylesXml: styleXml }))
    const table = doc.blocks[0].table!
    expect(table.borders?.top).toMatchObject({ style: 'single', szEighths: 16, color: '00FF00' })
    expect(table.borders?.insideH).toMatchObject({ style: 'dashed' })
    expect(table.cellMarTwips).toEqual({ left: 200, right: 200 })
  })
})

describe('cell alignment aggregation and trailing empty paragraphs', () => {
  it('sets cell.align only when every text paragraph declares the same jc', async () => {
    const xml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>' +
      // centered title + jc-less body: a td-level center would leak onto the body
      '<w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>要点</w:t></w:r></w:p></w:tc>' +
      // all text paragraphs agree (the empty one does not vote)
      '<w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>甲</w:t></w:r></w:p>' +
      '<w:p/>' +
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>乙</w:t></w:r></w:p></w:tc>' +
      '</w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const [mixed, agreed] = doc.blocks[0].table!.rows[0]
    expect(mixed.align).toBeUndefined()
    expect(agreed.align).toBe('center')
  })

  it('keeps trailing empty cell paragraphs and their empty-run size', async () => {
    const xml =
      '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>' +
      '<w:p><w:r><w:t>内容</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>' +
      '</w:tc></w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.paras).toEqual(['内容', ''])
    expect(cell.richParas?.[1].emptyRunSizeHalfPoints).toBe(2)
    expect(cell.richParas?.[1].runs).toEqual([])
  })
})
