import { describe, expect, it } from 'vitest'
import { generateTableModelXml, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

function tc(text: string, tcPr = ''): string {
  const pr = tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''
  return `<w:tc>${pr}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
}

function gridXml(widths: number[]): string {
  return `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`
}

async function tableOf(bodyXml: string) {
  const doc = await parseDocx(await buildDocx({ bodyXml }))
  const block = doc.blocks[0]
  expect(block.type).toBe('table')
  return block.table!
}

/** twips width a cell spans, from the resolved grid */
function cellWidth(widths: number[], row: { colSpan?: number }[], index: number): number {
  let col = 0
  for (let i = 0; i < index; i++) col += row[i].colSpan ?? 1
  return widths.slice(col, col + (row[index].colSpan ?? 1)).reduce((a, b) => a + b, 0)
}

describe('grid reconciliation from tcW boundaries', () => {
  it('rebuilds the grid when row spans under-cover it (govdocs roster shape)', async () => {
    // 7-column grid, but every row only spans 5 columns; the second cell would
    // otherwise land on the 28-twip sliver and collapse to one char per line
    const table = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="9692" w:type="dxa"/></w:tblPr>' +
        gridXml([79, 489, 4472, 810, 28, 3780, 34]) +
        '<w:tr>' +
        tc('Crew Name:', '<w:tcW w:w="5799" w:type="dxa"/><w:gridSpan w:val="4"/>') +
        tc('Host Unit/Address:', '<w:tcW w:w="3780" w:type="dxa"/>') +
        '</w:tr>' +
        '<w:tr>' +
        tc('a', '<w:tcW w:w="568" w:type="dxa"/><w:gridSpan w:val="2"/>') +
        tc('b', '<w:tcW w:w="4472" w:type="dxa"/>') +
        tc('c', '<w:tcW w:w="810" w:type="dxa"/>') +
        tc('d', '<w:tcW w:w="3842" w:type="dxa"/><w:gridSpan w:val="3"/>') +
        '</w:tr>' +
        '</w:tbl>',
    )
    const widths = table.colWidthsTwips!
    expect(cellWidth(widths, table.rows[0], 0)).toBe(5799)
    expect(cellWidth(widths, table.rows[0], 1)).toBe(3780)
    expect(cellWidth(widths, table.rows[1], 3)).toBe(3842)
    // no column may be a sliver that a text cell sits on alone
    expect(
      Math.min(...table.rows.flatMap((r) => r.map((_c, ci) => cellWidth(widths, r, ci)))),
    ).toBeGreaterThan(500)
  })

  it('re-aligns rows that dropped a gridSpan (govdocs workload shape)', async () => {
    // 8-column grid; the second data row lost the gridSpan on its label cell,
    // so its 7 cells drifted one column left and the label collapsed to 5 twips
    const auto = '<w:tcW w:w="0" w:type="auto"/>'
    const dataCells = tc('1,582', auto) + tc('1,537', auto) + tc('1,139', auto) + tc('1,093', auto)
    const table = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="10115" w:type="dxa"/></w:tblPr>' +
        gridXml([5, 3282, 1125, 1116, 1116, 1116, 1123, 1232]) +
        '<w:tr>' +
        tc('Labels', '<w:tcW w:w="3283" w:type="dxa"/><w:gridSpan w:val="2"/>') +
        tc('FY', '<w:tcW w:w="1124" w:type="dxa"/>') +
        dataCells +
        tc('Est.', '<w:tcW w:w="1230" w:type="dxa"/>') +
        '</w:tr>' +
        '<w:tr>' +
        tc('Pending', '<w:tcW w:w="3283" w:type="dxa"/>') +
        tc('260', '<w:tcW w:w="1124" w:type="dxa"/>') +
        dataCells +
        tc('89', '<w:tcW w:w="1230" w:type="dxa"/>') +
        '</w:tr>' +
        '</w:tbl>',
    )
    const widths = table.colWidthsTwips!
    // both label cells resolve to the same wide column(s)
    expect(cellWidth(widths, table.rows[0], 0)).toBe(3283)
    expect(cellWidth(widths, table.rows[1], 0)).toBe(3283)
    // both rows cover the same total width (no one-column drift / empty tail)
    const total = (r: { colSpan?: number }[]) =>
      r.reduce((sum, _c, i) => sum + cellWidth(widths, r, i), 0)
    expect(total(table.rows[1])).toBe(total(table.rows[0]))
  })

  it('leaves consistent grids untouched', async () => {
    const table = await tableOf(
      '<w:tbl>' +
        gridXml([2000, 4000, 2000]) +
        '<w:tr>' +
        tc('a', '<w:tcW w:w="2000" w:type="dxa"/>') +
        tc('bc', '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/>') +
        '</w:tr>' +
        '<w:tr>' +
        tc('a', '<w:tcW w:w="2000" w:type="dxa"/>') +
        tc('b', '<w:tcW w:w="4000" w:type="dxa"/>') +
        tc('c', '<w:tcW w:w="2000" w:type="dxa"/>') +
        '</w:tr>' +
        '</w:tbl>',
    )
    expect(table.colWidthsTwips).toEqual([2000, 4000, 2000])
    expect(table.rows[0][1].colSpan).toBe(2)
    expect(table.rows[1].every((c) => (c.colSpan ?? 1) === 1)).toBe(true)
  })
})

describe('gridBefore/gridAfter', () => {
  const STAGGERED =
    '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>' +
    gridXml([1000, 1000, 1000, 1000]) +
    '<w:tr><w:trPr><w:gridAfter w:val="1"/><w:wAfter w:w="1000" w:type="dxa"/></w:trPr>' +
    tc('a1', '<w:tcW w:w="1000" w:type="dxa"/>') +
    tc('a2', '<w:tcW w:w="1000" w:type="dxa"/>') +
    tc('a3', '<w:tcW w:w="1000" w:type="dxa"/>') +
    '</w:tr>' +
    '<w:tr><w:trPr><w:gridBefore w:val="2"/><w:wBefore w:w="2000" w:type="dxa"/></w:trPr>' +
    tc('b1', '<w:tcW w:w="1000" w:type="dxa"/>') +
    tc('b2', '<w:tcW w:w="1000" w:type="dxa"/>') +
    '</w:tr>' +
    '</w:tbl>'

  it('inserts borderless placeholder cells for the skipped columns', async () => {
    const table = await tableOf(STAGGERED)
    expect(table.colWidthsTwips).toEqual([1000, 1000, 1000, 1000])
    const [row1, row2] = table.rows
    expect(row1.map((c) => c.gridGap ?? false)).toEqual([false, false, false, true])
    expect(row2.map((c) => c.gridGap ?? false)).toEqual([true, false, false])
    expect(row2[0].colSpan).toBe(2)
    // offset row's first real cell starts at grid column 2
    expect(row2[1].paras[0]).toBe('b1')
    // rawTcPr still lands on the real cells despite the inserted placeholders
    expect(row2[0].rawTcPr).toBeUndefined()
    expect(row2[1].rawTcPr).toContain('<w:tcW w:w="1000"')
    expect(row1[3].rawTcPr).toBeUndefined()
  })

  it('regenerates placeholders as trPr gridBefore/gridAfter, not cells', async () => {
    const table = await tableOf(STAGGERED)
    const xml = generateTableModelXml(table)
    expect(xml.match(/<w:tc>/g)).toHaveLength(5)
    // rawTrPr passthrough keeps the original offsets exactly once
    expect(xml.match(/<w:gridBefore\b/g)).toHaveLength(1)
    expect(xml.match(/<w:gridAfter\b/g)).toHaveLength(1)
    expect(xml).toContain('<w:wBefore w:w="2000"')
    const bare = { ...table, rawTrPrs: undefined }
    const regen = generateTableModelXml(bare)
    expect(regen).toContain('<w:gridBefore w:val="2"/>')
    expect(regen).toContain('<w:wBefore w:w="2000" w:type="dxa"/>')
    expect(regen.match(/<w:tc>/g)).toHaveLength(5)
  })
})

describe('cell spacing and shading patterns', () => {
  it('parses w:tblCellSpacing and table-level shading', async () => {
    const table = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' +
        '<w:tblCellSpacing w:w="142" w:type="dxa"/>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="FFC000"/></w:tblPr>' +
        gridXml([2000, 2000]) +
        '<w:tr>' +
        tc('a', '<w:tcW w:w="2000" w:type="dxa"/>') +
        tc('b', '<w:tcW w:w="2000" w:type="dxa"/>') +
        '</w:tr>' +
        '</w:tbl>',
    )
    expect(table.cellSpacingTwips).toBe(142)
    expect(table.fill).toBe('FFC000')
  })

  it('approximates pattern shading as blended fills', async () => {
    const table = await tableOf(
      '<w:tbl>' +
        gridXml([2000, 2000, 2000]) +
        '<w:tr>' +
        tc('gray', '<w:shd w:val="pct5" w:color="auto" w:fill="auto"/>') +
        tc('stripe', '<w:shd w:val="reverseDiagStripe" w:color="6FB31A" w:fill="auto"/>') +
        tc('plain', '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>') +
        '</w:tr>' +
        '</w:tbl>',
    )
    const [gray, stripe, plain] = table.rows[0]
    expect(gray.fill).toBe('F2F2F2')
    expect(stripe.fill).toBe('B7D98D')
    expect(plain.fill).toBeUndefined()
  })
})
