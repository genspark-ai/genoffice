import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  parseDocx,
  saveDocx,
} from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { serializeRangeToHtml } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Guard tests for the table half of issue #175: the model can only rewrite a
 * table by echoing its <table> HTML through replace_blocks, and that HTML is
 * plain cell text — no widths, borders, shading, fonts. The rewritten table now
 * inherits every table/row/cell property from the table it replaces, and cells
 * whose text did not change keep their content verbatim.
 */

const NUM_IDS = { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID }
const TRACK = { author: 'AI Assistant' }

const HEADER_RPR =
  '<w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri"/><w:b/><w:color w:val="1F4E78"/></w:rPr>'
const BODY_RPR =
  '<w:rPr><w:rFonts w:ascii="Cambria" w:eastAsia="Cambria"/><w:color w:val="595959"/><w:sz w:val="20"/></w:rPr>'
const TOTAL_RPR = '<w:rPr><w:b/><w:sz w:val="20"/></w:rPr>'

const headerCell = (text: string) =>
  `<w:tc><w:tcPr><w:shd w:val="clear" w:fill="D9EAF7"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${HEADER_RPR}<w:t>${text}</w:t></w:r></w:p></w:tc>`
const cell = (text: string, rPr: string, jc?: string) =>
  `<w:tc><w:p>${jc ? `<w:pPr><w:jc w:val="${jc}"/></w:pPr>` : ''}<w:r>${rPr}<w:t>${text}</w:t></w:r></w:p></w:tc>`

/** 2 unequal columns; shaded bold header; Cambria 10pt body; bold total row; a paragraph after */
const STYLED_TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>' +
  `<w:tr><w:trPr><w:trHeight w:val="500"/></w:trPr>${headerCell('Metric')}${headerCell('Value')}</w:tr>` +
  `<w:tr>${cell('Revenue', BODY_RPR)}${cell('100', BODY_RPR, 'right')}</w:tr>` +
  `<w:tr>${cell('Total', TOTAL_RPR)}${cell('100', TOTAL_RPR, 'right')}</w:tr></w:tbl>` +
  '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>After the table.</w:t></w:r></w:p>'

/** header row shaded but not bold */
const FILL_ONLY_HEADER =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr><w:p><w:r><w:t>Role</w:t></w:r></w:p></w:tc></w:tr>' +
  `<w:tr>${cell('Ada', BODY_RPR)}${cell('Engineer', BODY_RPR)}</w:tr></w:tbl>`

/** A spans two rows (vMerge); the model sees an empty second-row cell under it */
const MERGED_TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/><w:shd w:val="clear" w:fill="FFF2CC"/></w:tcPr><w:p><w:r><w:t>Group</w:t></w:r></w:p></w:tc>' +
  `${cell('First', BODY_RPR)}</w:tr>` +
  '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>' +
  `${cell('Second', BODY_RPR)}</w:tr></w:tbl>`

/** a title cell spanning both columns (gridSpan) above two body rows */
const HMERGED_TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:t>Quarter</w:t></w:r></w:p></w:tc></w:tr>' +
  `<w:tr>${cell('Revenue', BODY_RPR)}${cell('100', BODY_RPR)}</w:tr>` +
  `<w:tr>${cell('Cost', BODY_RPR)}${cell('60', BODY_RPR)}</w:tr></w:tbl>`

/** a spanning title row above a vertically merged group column */
const BOTH_MERGED_TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:t>Quarter</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/><w:shd w:val="clear" w:fill="FFF2CC"/></w:tcPr><w:p><w:r><w:t>Group</w:t></w:r></w:p></w:tc>' +
  `${cell('First', BODY_RPR)}</w:tr>` +
  `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>${cell('Second', BODY_RPR)}</w:tr></w:tbl>`

/** a bulleted list inside a cell (numbering part from the build helper: numId 1 = bullet) */
const bulletPara = (text: string) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r>${BODY_RPR}<w:t>${text}</w:t></w:r></w:p>`
const LIST_CELL_TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr>${cell('Topic', BODY_RPR)}<w:tc>${bulletPara('First point')}${bulletPara('Second point')}</w:tc></w:tr></w:tbl>`

type Json = {
  type: string
  attrs: Record<string, unknown>
  content?: Json[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

async function open(bodyXml: string, withNumbering = false) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withNumbering }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

async function replaceBlocks(
  editor: Editor,
  start: number,
  end: number,
  html: string,
  track?: typeof TRACK,
) {
  const r = await executeTool(
    editor,
    {
      id: 't',
      name: 'replace_blocks',
      input: { startBlockIndex: start, endBlockIndex: end, html },
    },
    NUM_IDS,
    track,
  )
  expect(r.isError).toBeFalsy()
  return r
}

const blocks = (editor: Editor) => ((editor.getJSON() as PmNode).content ?? []) as Json[]
const cellAt = (table: Json, r: number, c: number) => table.content![r].content![c]
const cellRuns = (table: Json, r: number, c: number) =>
  cellAt(table, r, c).content![0].content ?? []
const cellText = (table: Json, r: number, c: number) =>
  cellRuns(table, r, c)
    .map((n) => n.text ?? '')
    .join('')
const textStyle = (run: Json) => run.marks?.find((m) => m.type === 'docTextStyle')?.attrs
const hasMark = (run: Json, type: string) => (run.marks ?? []).some((m) => m.type === type)

async function saveAndReparse(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
  return reparsed.blocks.find((b) => b.type === 'table')!.table!
}

describe('replace_blocks keeps a rewritten table formatted like the table it replaces', () => {
  it('the model only ever sees plain cell text', async () => {
    const { editor } = await open(STYLED_TABLE)
    const html = serializeRangeToHtml(editor, 0, 0)
    expect(html).toBe(
      '<table><tr><th>Metric</th><th>Value</th></tr><tr><td>Revenue</td><td>100</td></tr><tr><td>Total</td><td>100</td></tr></table>',
    )
    expect(html).not.toMatch(/D9EAF7|Calibri|Cambria|3000|center|right/)
  })

  it('translating the cells keeps table, row and cell properties; unchanged cells stay verbatim', async () => {
    const { editor, parsed } = await open(STYLED_TABLE)
    const [before] = blocks(editor)
    const untouchedCell = JSON.stringify(cellAt(before, 1, 1))

    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><th>指标</th><th>数值</th></tr><tr><td>收入</td><td>100</td></tr><tr><td>合计</td><td>100</td></tr></table>',
    )

    const [table, after] = blocks(editor)
    expect(blocks(editor)).toHaveLength(2)
    expect(after.type).toBe('docParagraph')
    // table props: the anchor, the unequal column grid, the table style
    expect(table.type).toBe('docTable')
    expect(table.attrs.docxIndex).toBe(before.attrs.docxIndex)
    expect(table.attrs.colWidthsPct).toEqual(before.attrs.colWidthsPct)
    expect(table.attrs.colWidthsPct).toEqual([37.5, 62.5])
    expect(table.attrs.tblStyleId).toBe('TableGrid')
    expect(table.attrs.originalStructure).toBe(before.attrs.originalStructure)
    // row props
    expect(table.content![0].attrs.heightTwips).toBe(500)
    // header cell: shading, header node type, column width, centered paragraph,
    // and the rewritten text in the header run style
    const header = cellAt(table, 0, 0)
    expect(header.type).toBe('docTableHeader')
    expect(header.attrs).toMatchObject({ fill: 'D9EAF7', colwidth: [200] })
    expect(header.content![0].attrs.align).toBe('center')
    expect(cellText(table, 0, 0)).toBe('指标')
    expect(cellText(table, 0, 1)).toBe('数值')
    const headerRun = cellRuns(table, 0, 0)[0]
    expect(hasMark(headerRun, 'bold')).toBe(true)
    expect(textStyle(headerRun)).toMatchObject({ font: 'Calibri', color: '1F4E78' })
    // body cell: the Cambria 10pt gray run style, not bold
    const body = cellRuns(table, 1, 0)[0]
    expect(body.text).toBe('收入')
    expect(hasMark(body, 'bold')).toBe(false)
    expect(textStyle(body)).toMatchObject({
      font: 'Cambria',
      fontAscii: 'Cambria',
      sizeHalfPoints: 20,
      color: '595959',
    })
    // total cell: bold inherited from the old run even though <td> carries no <strong>
    const total = cellRuns(table, 2, 0)[0]
    expect(total.text).toBe('合计')
    expect(hasMark(total, 'bold')).toBe(true)
    expect(textStyle(total)).toMatchObject({ sizeHalfPoints: 20 })
    // a cell whose text did not change is byte-for-byte the old cell (rawRPr included)
    expect(JSON.stringify(cellAt(table, 1, 1))).toBe(untouchedCell)
    expect(cellAt(table, 2, 1).content![0].attrs.align).toBe('right')

    const saved = await saveAndReparse(editor, parsed)
    expect(saved.colWidthsTwips).toEqual([3000, 5000])
    expect(saved.tblStyleId).toBe('TableGrid')
    expect(saved.rows.map((row) => row.map((c) => c.paras))).toEqual([
      [['指标'], ['数值']],
      [['收入'], ['100']],
      [['合计'], ['100']],
    ])
    expect(saved.rows[0][0]).toMatchObject({ fill: 'D9EAF7', bold: true })
    expect(saved.rows[1][0].richParas?.[0].runs[0]).toMatchObject({
      text: '收入',
      font: 'Cambria',
      sizeHalfPoints: 20,
      color: '595959',
    })
    expect(saved.rows[2][0].richParas?.[0].runs[0]).toMatchObject({ text: '合计', bold: true })
    expect(saved.rows[2][1].richParas?.[0].align).toBe('right')
  })

  it('a shaded header row that was not bold does not turn bold because the model echoed <th>', async () => {
    const { editor } = await open(FILL_ONLY_HEADER)
    expect(serializeRangeToHtml(editor, 0, 0)).toContain('<th>Name</th>')
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><th>Full name</th><th>Role</th></tr><tr><td>Ada Lovelace</td><td>Engineer</td></tr></table>',
    )
    const [table] = blocks(editor)
    const header = cellAt(table, 0, 0)
    expect(header.attrs).toMatchObject({ fill: 'EEEEEE', bold: false })
    expect(cellText(table, 0, 0)).toBe('Full name')
    expect(hasMark(cellRuns(table, 0, 0)[0], 'bold')).toBe(false)
    expect(textStyle(cellRuns(table, 1, 0)[0])).toMatchObject({ font: 'Cambria' })
  })

  it('a row the model adds is formatted like the row above it; the other rows stay verbatim', async () => {
    const { editor, parsed } = await open(STYLED_TABLE)
    const [before] = blocks(editor)
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><th>Metric</th><th>Value</th></tr><tr><td>Revenue</td><td>100</td></tr><tr><td>Cost</td><td>60</td></tr><tr><td>Total</td><td>40</td></tr></table>',
    )
    const [table] = blocks(editor)
    expect(table.content).toHaveLength(4)
    expect(table.attrs.docxIndex).toBe(before.attrs.docxIndex)
    expect(table.attrs.colWidthsPct).toEqual([37.5, 62.5])
    expect(table.attrs.tblStyleId).toBe('TableGrid')
    expect(JSON.stringify(table.content![0])).toBe(JSON.stringify(before.content![0]))
    expect(JSON.stringify(table.content![1])).toBe(JSON.stringify(before.content![1]))
    // the new Cost row: body cell type, the old column widths, Cambria 10pt, right-aligned value
    const cost = table.content![2]
    expect(cost.content!.map((c) => c.type)).toEqual(['docTableCell', 'docTableCell'])
    expect(cost.content!.map((c) => c.attrs.colwidth)).toEqual([[200], [333]])
    expect(cellText(table, 2, 0)).toBe('Cost')
    expect(textStyle(cellRuns(table, 2, 0)[0])).toMatchObject({
      font: 'Cambria',
      sizeHalfPoints: 20,
    })
    expect(hasMark(cellRuns(table, 2, 0)[0], 'bold')).toBe(false)
    expect(cellAt(table, 2, 1).content![0].attrs.align).toBe('right')
    // the total row keeps its bold and its new value
    expect(hasMark(cellRuns(table, 3, 0)[0], 'bold')).toBe(true)
    expect(cellText(table, 3, 1)).toBe('40')

    const saved = await saveAndReparse(editor, parsed)
    expect(saved.rows).toHaveLength(4)
    // a structural change regenerates the table from the (px-rounded) model grid
    expect(saved.colWidthsTwips?.map((w) => Math.round(w / 20))).toEqual([150, 250])
    expect(saved.tblStyleId).toBe('TableGrid')
    expect(saved.rows[0][0]).toMatchObject({ fill: 'D9EAF7', bold: true })
    expect(saved.rows[2].map((c) => c.paras)).toEqual([['Cost'], ['60']])
    expect(saved.rows[2][0].richParas?.[0].runs[0]).toMatchObject({
      font: 'Cambria',
      sizeHalfPoints: 20,
    })
  })

  it('a column the model adds takes the look of its neighbour; the grid becomes an equal split', async () => {
    const { editor, parsed } = await open(STYLED_TABLE)
    const [before] = blocks(editor)
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><th>Metric</th><th>Value</th><th>Change</th></tr><tr><td>Revenue</td><td>100</td><td>+5%</td></tr><tr><td>Total</td><td>100</td><td>+5%</td></tr></table>',
    )
    const [table] = blocks(editor)
    expect(table.attrs.docxIndex).toBe(before.attrs.docxIndex)
    expect(table.attrs.tblStyleId).toBe('TableGrid')
    expect(table.content![0].attrs.heightTwips).toBe(500)
    expect((table.attrs.colWidthsPct as number[]).map((w) => Math.round(w))).toEqual([33, 33, 33])
    for (let c = 0; c < 3; c++) {
      expect(cellAt(table, 0, c).attrs).toMatchObject({ fill: 'D9EAF7', colspan: 1, rowspan: 1 })
      expect(cellAt(table, 0, c).type).toBe('docTableHeader')
    }
    expect(cellText(table, 0, 2)).toBe('Change')
    expect(hasMark(cellRuns(table, 0, 2)[0], 'bold')).toBe(true)
    expect(textStyle(cellRuns(table, 1, 2)[0])).toMatchObject({
      font: 'Cambria',
      sizeHalfPoints: 20,
    })
    expect(hasMark(cellRuns(table, 2, 2)[0], 'bold')).toBe(true)

    const saved = await saveAndReparse(editor, parsed)
    expect(saved.rows[0].map((c) => c.paras)).toEqual([['Metric'], ['Value'], ['Change']])
    expect(saved.rows[0][2]).toMatchObject({ fill: 'D9EAF7', bold: true })
    expect(saved.rows[1][2].richParas?.[0].runs[0]).toMatchObject({ text: '+5%', font: 'Cambria' })
  })

  it('merged cells survive: the empty cell the model echoes under a vertical merge is not a new cell', async () => {
    const { editor, parsed } = await open(MERGED_TABLE)
    expect(serializeRangeToHtml(editor, 0, 0)).toBe(
      '<table><tr><td>Group</td><td>First</td></tr><tr><td></td><td>Second</td></tr></table>',
    )
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><td>Category</td><td>1st</td></tr><tr><td></td><td>2nd</td></tr></table>',
    )
    const [table] = blocks(editor)
    expect(table.content![0].content).toHaveLength(2)
    expect(table.content![1].content).toHaveLength(1)
    expect(cellAt(table, 0, 0).attrs).toMatchObject({ rowspan: 2, fill: 'FFF2CC' })
    expect(cellText(table, 0, 0)).toBe('Category')
    expect(cellText(table, 1, 0)).toBe('2nd')

    const saved = await saveAndReparse(editor, parsed)
    expect(saved.rows[0][0]).toMatchObject({
      vMerge: 'restart',
      fill: 'FFF2CC',
      paras: ['Category'],
    })
    expect(saved.rows[1][0].vMerge).toBe('continue')
    expect(saved.rows[1][1].paras).toEqual(['2nd'])
  })

  it('adding a row below a horizontally merged title keeps the span, the grid and the widths', async () => {
    const { editor, parsed } = await open(HMERGED_TABLE)
    const [before] = blocks(editor)
    // the spanning cell is one <th>; the parser pads that row with an empty cell
    expect(serializeRangeToHtml(editor, 0, 0)).toBe(
      '<table><tr><th>Quarter</th></tr><tr><td>Revenue</td><td>100</td></tr><tr><td>Cost</td><td>60</td></tr></table>',
    )
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><th>Quarter</th></tr><tr><td>Revenue</td><td>100</td></tr><tr><td>Cost</td><td>60</td></tr><tr><td>Profit</td><td>40</td></tr></table>',
    )
    const [table] = blocks(editor)
    expect(table.content).toHaveLength(4)
    expect(table.attrs.colWidthsPct).toEqual(before.attrs.colWidthsPct)
    expect(table.attrs.colWidthsPct).toEqual([37.5, 62.5])
    // the title row is untouched: still one cell spanning both columns
    expect(table.content![0].content).toHaveLength(1)
    expect(cellAt(table, 0, 0).attrs).toMatchObject({ colspan: 2, fill: 'D9EAF7' })
    expect(JSON.stringify(table.content![0])).toBe(JSON.stringify(before.content![0]))
    // the new row takes the geometry and look of the row above it
    const profit = table.content![3]
    expect(profit.content!.map((c) => c.attrs.colspan)).toEqual([1, 1])
    expect(profit.content!.map((c) => c.attrs.colwidth)).toEqual([[200], [333]])
    expect(cellText(table, 3, 0)).toBe('Profit')
    expect(textStyle(cellRuns(table, 3, 0)[0])).toMatchObject({
      font: 'Cambria',
      sizeHalfPoints: 20,
    })

    const saved = await saveAndReparse(editor, parsed)
    expect(saved.rows).toHaveLength(4)
    expect(saved.rows[0]).toHaveLength(1)
    expect(saved.rows[0][0]).toMatchObject({ colSpan: 2, fill: 'D9EAF7', paras: ['Quarter'] })
    expect(saved.rows[3].map((c) => c.paras)).toEqual([['Profit'], ['40']])
  })

  it('adding a row to a vertically merged table dissolves the merge into plain cells instead of corrupting the grid', async () => {
    const { editor, parsed } = await open(MERGED_TABLE)
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><td>Group</td><td>First</td></tr><tr><td></td><td>Second</td></tr><tr><td></td><td>Third</td></tr></table>',
    )
    const [table] = blocks(editor)
    expect(table.content).toHaveLength(3)
    // every row has a cell in every column, no rowspan left to overlap the new row
    for (const row of table.content!) {
      expect(row.content).toHaveLength(2)
      expect(row.content!.map((c) => c.attrs.rowspan)).toEqual([1, 1])
    }
    expect(cellAt(table, 0, 0).attrs.fill).toBe('FFF2CC')
    expect(cellText(table, 0, 0)).toBe('Group')
    expect(cellText(table, 2, 1)).toBe('Third')
    expect(textStyle(cellRuns(table, 2, 1)[0])).toMatchObject({ font: 'Cambria' })
    expect(table.attrs.tblStyleId).toBe('TableGrid')

    const saved = await saveAndReparse(editor, parsed)
    expect(saved.rows).toHaveLength(3)
    expect(saved.rows.map((row) => row.map((c) => c.paras))).toEqual([
      [['Group'], ['First']],
      [[''], ['Second']],
      [[''], ['Third']],
    ])
    expect(saved.rows.every((row) => row.every((c) => !c.vMerge))).toBe(true)
  })

  it('only the rows a vertical merge runs through dissolve; a spanning title row elsewhere keeps its span', async () => {
    const { editor, parsed } = await open(BOTH_MERGED_TABLE)
    const [before] = blocks(editor)
    expect(serializeRangeToHtml(editor, 0, 0)).toBe(
      '<table><tr><th>Quarter</th></tr><tr><td>Group</td><td>First</td></tr><tr><td></td><td>Second</td></tr></table>',
    )
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><th>Quarter</th></tr><tr><td>Group</td><td>First</td></tr><tr><td></td><td>Second</td></tr><tr><td>Other</td><td>Third</td></tr></table>',
    )
    const [table] = blocks(editor)
    expect(table.content).toHaveLength(4)
    expect(table.attrs.colWidthsPct).toEqual(before.attrs.colWidthsPct)
    // the title row is not part of the vertical merge: byte-identical, span kept
    expect(table.content![0].content).toHaveLength(1)
    expect(cellAt(table, 0, 0).attrs).toMatchObject({ colspan: 2, rowspan: 1, fill: 'D9EAF7' })
    expect(JSON.stringify(table.content![0])).toBe(JSON.stringify(before.content![0]))
    // the merged group dissolves: the spanning cell drops its rowspan, the slot
    // below it becomes a plain cell in the same look
    expect(cellAt(table, 1, 0).attrs).toMatchObject({ colspan: 1, rowspan: 1, fill: 'FFF2CC' })
    expect(cellText(table, 1, 0)).toBe('Group')
    expect(table.content![2].content).toHaveLength(2)
    expect(cellAt(table, 2, 0).attrs).toMatchObject({ colspan: 1, rowspan: 1, fill: 'FFF2CC' })
    expect(cellAt(table, 2, 0).attrs.colwidth).toEqual(cellAt(table, 1, 0).attrs.colwidth)
    expect(cellText(table, 2, 0)).toBe('')
    expect(cellText(table, 2, 1)).toBe('Second')
    // the new row takes the look of the row above, with no merge to fall under
    expect(table.content![3].content!.map((c) => c.attrs.rowspan)).toEqual([1, 1])
    expect(cellText(table, 3, 0)).toBe('Other')
    expect(textStyle(cellRuns(table, 3, 1)[0])).toMatchObject({ font: 'Cambria' })

    const saved = await saveAndReparse(editor, parsed)
    expect(saved.rows.map((row) => row.map((c) => c.paras))).toEqual([
      [['Quarter']],
      [['Group'], ['First']],
      [[''], ['Second']],
      [['Other'], ['Third']],
    ])
    expect(saved.rows[0][0]).toMatchObject({ colSpan: 2, fill: 'D9EAF7' })
    expect(saved.rows[2][0].fill).toBe('FFF2CC')
    expect(saved.rows.every((row) => row.every((c) => !c.vMerge))).toBe(true)
  })

  it('a bulleted cell stays bulleted when the model rewrites its points', async () => {
    const { editor, parsed } = await open(LIST_CELL_TABLE, true)
    expect(serializeRangeToHtml(editor, 0, 0)).toContain('<td>First point<br>Second point</td>')
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><td>Topic</td><td>Point one<br>Point two<br>Point three</td></tr></table>',
    )
    const [table] = blocks(editor)
    const items = cellAt(table, 0, 1).content!
    expect(items.map((n) => n.type)).toEqual(['docListItem', 'docListItem', 'docListItem'])
    expect(items.map((n) => n.content![0].text)).toEqual(['Point one', 'Point two', 'Point three'])
    expect(new Set(items.map((n) => n.attrs.numId)).size).toBe(1)
    expect(textStyle(items[2].content![0])).toMatchObject({ font: 'Cambria', sizeHalfPoints: 20 })

    const saved = await saveAndReparse(editor, parsed)
    expect(saved.rows[0][1].paras).toEqual(['Point one', 'Point two', 'Point three'])
    expect(saved.rows[0][1].richParas?.map((p) => p.list?.numId)).toEqual(['1', '1', '1'])
  })

  it('a tracked rewrite inherits the formatting but leaves the anchor with the struck-through table', async () => {
    const { editor } = await open(STYLED_TABLE)
    const [before] = blocks(editor)
    await replaceBlocks(
      editor,
      0,
      0,
      '<table><tr><th>指标</th><th>数值</th></tr><tr><td>收入</td><td>100</td></tr><tr><td>合计</td><td>100</td></tr></table>',
      TRACK,
    )
    const [old, table] = blocks(editor)
    expect(old.attrs.blockRevision).toMatchObject({ kind: 'del' })
    expect(old.attrs.docxIndex).toBe(before.attrs.docxIndex)
    expect(table.attrs.blockRevision).toMatchObject({ kind: 'ins' })
    expect(table.attrs.docxIndex).toBeNull()
    expect(table.attrs.colWidthsPct).toEqual([37.5, 62.5])
    expect(cellAt(table, 0, 0).attrs.fill).toBe('D9EAF7')
    expect(textStyle(cellRuns(table, 1, 0)[0])).toMatchObject({ font: 'Cambria' })
  })

  it('a range mixing a table and a paragraph formats each from its own kind of block', async () => {
    const { editor } = await open(STYLED_TABLE)
    const [beforeTable, beforePara] = blocks(editor)
    await replaceBlocks(
      editor,
      0,
      1,
      '<table><tr><th>Metric</th><th>Value</th></tr><tr><td>Revenue</td><td>120</td></tr><tr><td>Total</td><td>120</td></tr></table><p>Below the table.</p>',
    )
    const [table, para] = blocks(editor)
    expect(table.attrs.docxIndex).toBe(beforeTable.attrs.docxIndex)
    expect(cellText(table, 1, 1)).toBe('120')
    expect(textStyle(cellRuns(table, 1, 1)[0])).toMatchObject({ font: 'Cambria' })
    expect(para.attrs.docxIndex).toBe(beforePara.attrs.docxIndex)
    expect(para.content![0].text).toBe('Below the table.')
  })
})
