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

async function tableOf(bodyXml: string, sectPrExtra?: string) {
  const doc = await parseDocx(await buildDocx({ bodyXml, sectPrExtra }))
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

  it("a table with no declared width takes its cells' preferred widths over the grid", async () => {
    const cell = tc('title', '<w:tcW w:w="6480" w:type="dxa"/>')
    const auto = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/></w:tblPr>' +
        gridXml([10469]) +
        `<w:tr>${cell}</w:tr></w:tbl>`,
    )
    expect(auto.colWidthsTwips).toEqual([6480])
    // a declared dxa width keeps the (consistent-ratio) grid
    const declared = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="10469" w:type="dxa"/></w:tblPr>' +
        gridXml([10469]) +
        `<w:tr>${cell}</w:tr></w:tbl>`,
    )
    expect(declared.colWidthsTwips).toEqual([10469])
    // a percent width is declared too: the grid stays authoritative
    const pct = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="100%" w:type="pct"/></w:tblPr>' +
        gridXml([10469]) +
        `<w:tr>${cell}</w:tr></w:tbl>`,
    )
    expect(pct.colWidthsTwips).toEqual([10469])
  })

  it('keeps the grid of an undeclared-width table whose preferred widths overflow the column', async () => {
    // Word shrinks over-wide tcW (9740 under a 9026 column) to the column and saves
    // that layout as the grid; the same ratios in tcW must not widen it again
    const row =
      '<w:tr>' +
      tc('a', '<w:tcW w:w="3840" w:type="dxa"/>') +
      tc('b', '<w:tcW w:w="1520" w:type="dxa"/>') +
      tc('c', '<w:tcW w:w="4380" w:type="dxa"/>') +
      '</w:tr>'
    const table = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        gridXml([3554, 1411, 4051]) +
        row +
        row +
        '</w:tbl>',
    )
    expect(table.colWidthsTwips).toEqual([3554, 1411, 4051])
  })

  it('a same-ratio grid well below the column is still a stale placeholder when tcW overflow', async () => {
    const row =
      '<w:tr>' +
      tc('a', '<w:tcW w:w="4000" w:type="dxa"/>') +
      tc('b', '<w:tcW w:w="4000" w:type="dxa"/>') +
      tc('c', '<w:tcW w:w="4000" w:type="dxa"/>') +
      '</w:tr>'
    const table = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        gridXml([2000, 2000, 2000]) +
        row +
        '</w:tbl>',
    )
    expect(table.colWidthsTwips).toEqual([4000, 4000, 4000])
  })

  it('in a two-column section the shrunk layout spans one newspaper column', async () => {
    // (9026 - 720) / 2 = 4153 per column; tcW 5000 overflow it
    const row =
      '<w:tr>' +
      tc('a', '<w:tcW w:w="2000" w:type="dxa"/>') +
      tc('b', '<w:tcW w:w="1000" w:type="dxa"/>') +
      tc('c', '<w:tcW w:w="2000" w:type="dxa"/>') +
      '</w:tr>'
    const xml =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      gridXml([1661, 830, 1662]) +
      row +
      '</w:tbl>'
    expect((await tableOf(xml, '<w:cols w:num="2" w:space="720"/>')).colWidthsTwips).toEqual([
      1661, 830, 1662,
    ])
    // the same grid under a single column is a stale placeholder
    expect((await tableOf(xml)).colWidthsTwips).toEqual([2000, 1000, 2000])
  })

  it('a positive table indent narrows the width the shrunk layout spans', async () => {
    // 9026 - 720 indent = 8306; tcW 10000 overflow the column
    const row =
      '<w:tr>' +
      tc('a', '<w:tcW w:w="4000" w:type="dxa"/>') +
      tc('b', '<w:tcW w:w="2000" w:type="dxa"/>') +
      tc('c', '<w:tcW w:w="4000" w:type="dxa"/>') +
      '</w:tr>'
    const table = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="720" w:type="dxa"/></w:tblPr>' +
        gridXml([3322, 1661, 3323]) +
        row +
        '</w:tbl>',
    )
    expect(table.colWidthsTwips).toEqual([3322, 1661, 3323])
  })

  it('a negative table indent widens the box the shrunk layout spans', async () => {
    // 9026 + 500 = 9526 (the box grows to the left, the right edge stays on the
    // margin); tcW 10600 overflow it, the saved unequal grid at 9526 is the layout
    const row =
      '<w:tr>' +
      tc('a', '<w:tcW w:w="4200" w:type="dxa"/>') +
      tc('b', '<w:tcW w:w="2100" w:type="dxa"/>') +
      tc('c', '<w:tcW w:w="4300" w:type="dxa"/>') +
      '</w:tr>'
    const tbl = (ind: string) =>
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${ind}</w:tblPr>` +
      gridXml([3800, 1900, 3826]) +
      row +
      '</w:tbl>'
    const widened = await tableOf(tbl('<w:tblInd w:w="-500" w:type="dxa"/>'))
    expect(widened.colWidthsTwips).toEqual([3800, 1900, 3826])
    expect(widened.layoutGrid).toBe(true)
    // without the indent the same grid overshoots the column: a stale placeholder
    const plain = await tableOf(tbl(''))
    expect(plain.colWidthsTwips).toEqual([4200, 2100, 4300])
    expect(plain.layoutGrid).toBeUndefined()
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

describe('tblW-auto layout grid flag', () => {
  const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
  const settingsPart = (compat: number) => ({
    path: 'word/settings.xml',
    xml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>' +
      `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${compat}"/>` +
      '</w:compat></w:settings>',
    contentType: SETTINGS_CT,
  })
  const row = (tcws: Array<number | null>) =>
    '<w:tr>' +
    tcws
      .map((w, i) =>
        tc(
          'c' + i,
          w === null ? '<w:tcW w:w="0" w:type="auto"/>' : `<w:tcW w:w="${w}" w:type="dxa"/>`,
        ),
      )
      .join('') +
    '</w:tr>'

  it('marks the unequal grid an autofit table keeps as Word-laid-out', async () => {
    // audit table: grid 9016 = column, tcW 9740 overflow -> the grid is Word's shrunk layout
    const kept = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        gridXml([3554, 1411, 4051]) +
        row([3840, 1520, 4380]) +
        '</w:tbl>',
    )
    expect(kept.colWidthsTwips).toEqual([3554, 1411, 4051])
    expect(kept.layoutGrid).toBe(true)
    // tcW auto: Word re-autofits from content and saves the result as the grid
    const autoCells = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        gridXml([3813, 2508, 929, 2110]) +
        row([null, null, null, null]) +
        '</w:tbl>',
    )
    expect(autoCells.colWidthsTwips).toEqual([3813, 2508, 929, 2110])
    expect(autoCells.layoutGrid).toBe(true)
  })

  it('leaves placeholder grids, tcW-driven widths and declared-width tables unflagged', async () => {
    const placeholder = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        gridXml([2000, 2000, 2000]) +
        row([3000, 3000, 3000]) +
        '</w:tbl>',
    )
    expect(placeholder.colWidthsTwips).toEqual([3000, 3000, 3000])
    expect(placeholder.layoutGrid).toBeUndefined()
    const declared = await tableOf(
      '<w:tbl><w:tblPr><w:tblW w:w="9016" w:type="dxa"/></w:tblPr>' +
        gridXml([3554, 1411, 4051]) +
        row([3554, 1411, 4051]) +
        '</w:tbl>',
    )
    expect(declared.layoutGrid).toBeUndefined()
  })

  it('legacy compat lets the shrunk layout hang by the cell margins', async () => {
    // 9026 column - 1000 indent = 8026 of cell text; compat 14 borders hang by the
    // 200-twip side margins -> a saved grid of 8426 is that layout, compat 15 has no
    // hang and the same grid is a stale placeholder under the overflowing tcW
    const xml =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="1000" w:type="dxa"/>' +
      '<w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      gridXml([1200, 3400, 1600, 2226]) +
      row([1300, 3600, 1700, 2400]) +
      '</w:tbl>'
    const of = async (compat: number) =>
      (await parseDocx(await buildDocx({ bodyXml: xml, extraParts: [settingsPart(compat)] })))
        .blocks[0].table!
    const legacy = await of(14)
    expect(legacy.colWidthsTwips).toEqual([1200, 3400, 1600, 2226])
    expect(legacy.layoutGrid).toBe(true)
    const modern = await of(15)
    expect(modern.colWidthsTwips).toEqual([1300, 3600, 1700, 2400])
    expect(modern.layoutGrid).toBeUndefined()
  })
})

describe('hostile colSpan values', () => {
  it('clamps non-finite and huge spans instead of throwing or emitting invalid OOXML', async () => {
    const start = Date.now()
    const xml = generateTableModelXml({
      rows: [
        [
          { paras: ['a'], colSpan: Infinity },
          { paras: ['b'], colSpan: 1e9 },
        ],
      ],
    })
    expect(Date.now() - start).toBeLessThan(5000)
    expect(xml).not.toContain('Infinity')
    expect(xml).toContain('<w:gridSpan w:val="1000"/>')
    // the clamped model round-trips through the parser with a finite grid
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].table).toBeDefined()
  })
})
