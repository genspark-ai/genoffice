import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeOps } from '../src/renderer/ai/ops'
import { describeTable, parseTableLength } from '../src/renderer/ai/table-ops'
import { pmTableToModel } from '../src/renderer/editor/convert'
import { setModuleLang } from '../src/renderer/i18n/locale'

setModuleLang('en')

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: t ? [{ type: 'text', text: t }] : [],
})

const cell = (t: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docTableCell',
  attrs: { colspan: 1, rowspan: 1, colwidth: [100], ...attrs },
  content: [para(t)],
})

/** 3×3 grid with px widths, header row shaded */
function table(attrs: Record<string, unknown> = {}): JsonNode {
  return {
    type: 'docTable',
    attrs: { docxIndex: 7, widthPx: 300, colWidthsPct: [33.3, 33.3, 33.4], ...attrs },
    content: [
      {
        type: 'docTableRow',
        content: [
          cell('A1', { fill: 'DDDDDD' }),
          cell('B1', { fill: 'DDDDDD' }),
          cell('C1', { fill: 'DDDDDD' }),
        ],
      },
      { type: 'docTableRow', content: [cell('A2'), cell('B2'), cell('C2')] },
      { type: 'docTableRow', content: [cell('A3'), cell('B3'), cell('C3')] },
    ],
  }
}

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

const doc = () => [para('Intro'), table(), para('Outro')]
const target = { blockIndexes: [1] }
const tableNode = (editor: Editor): PmNode => editor.state.doc.child(1)
const rowTexts = (editor: Editor): string[][] => {
  const out: string[][] = []
  tableNode(editor).forEach((row) => {
    const cells: string[] = []
    row.forEach((c) => cells.push(c.textContent))
    out.push(cells)
  })
  return out
}

describe('table structure ops', () => {
  it('inserts rows after/before and copies the neighbouring row formatting', () => {
    const editor = createEditor(doc())
    const r = executeOps(editor, [{ op: 'insertTableRow', target, row: 0, count: 2 }])
    expect(r.ok, r.error).toBe(true)
    expect(rowTexts(editor)).toEqual([
      ['A1', 'B1', 'C1'],
      ['', '', ''],
      ['', '', ''],
      ['A2', 'B2', 'C2'],
      ['A3', 'B3', 'C3'],
    ])
    // copied from row 0 (the row before the insertion point)
    expect(tableNode(editor).child(1).child(0).attrs.fill).toBe('DDDDDD')
    expect(tableNode(editor).child(1).child(0).attrs.colwidth).toEqual([100])

    const before = executeOps(editor, [
      { op: 'insertTableRow', target, row: 0, position: 'before' },
    ])
    expect(before.ok).toBe(true)
    expect(rowTexts(editor)[0]).toEqual(['', '', ''])
    expect(rowTexts(editor)[1]).toEqual(['A1', 'B1', 'C1'])
  })

  it('deletes rows and refuses to delete them all', () => {
    const editor = createEditor(doc())
    expect(executeOps(editor, [{ op: 'deleteTableRow', target, row: 1 }]).ok).toBe(true)
    expect(rowTexts(editor)).toEqual([
      ['A1', 'B1', 'C1'],
      ['A3', 'B3', 'C3'],
    ])
    const all = executeOps(editor, [{ op: 'deleteTableRow', target, row: 0, count: 2 }])
    expect(all.ok).toBe(false)
    expect(all.error).toContain('cannot delete every row')
    const range = executeOps(editor, [{ op: 'deleteTableRow', target, row: 5 }])
    expect(range.error).toContain('out of range (0-1)')
  })

  it('inserts and deletes columns keeping the px grid consistent', () => {
    const editor = createEditor(doc())
    const r = executeOps(editor, [{ op: 'insertTableColumn', target, col: 2 }])
    expect(r.ok, r.error).toBe(true)
    expect(rowTexts(editor)[0]).toEqual(['A1', 'B1', 'C1', ''])
    const t = tableNode(editor)
    expect(t.attrs.widthPx).toBe(400)
    expect((t.attrs.colWidthsPct as number[]).map((p) => Math.round(p))).toEqual([25, 25, 25, 25])
    t.forEach((row) => row.forEach((c) => expect(c.attrs.colwidth).toEqual([100])))
    expect(t.child(0).child(3).attrs.fill).toBe('DDDDDD')
    const model = pmTableToModel(editor.getJSON().content![1] as never)
    expect(model.colWidthsTwips).toEqual([1500, 1500, 1500, 1500])

    const before = executeOps(editor, [
      { op: 'insertTableColumn', target, col: 0, position: 'before' },
    ])
    expect(before.ok).toBe(true)
    expect(rowTexts(editor)[1]).toEqual(['', 'A2', 'B2', 'C2', ''])

    const del = executeOps(editor, [{ op: 'deleteTableColumn', target, col: 0, count: 2 }])
    expect(del.ok, del.error).toBe(true)
    expect(rowTexts(editor)[1]).toEqual(['B2', 'C2', ''])
    expect(tableNode(editor).attrs.widthPx).toBe(300)
    const last = executeOps(editor, [{ op: 'deleteTableColumn', target, col: 0, count: 3 }])
    expect(last.error).toContain('cannot delete every column')
  })

  it('refuses a column delete that would leave a rowspan-covered row without cells', () => {
    const editor = createEditor(doc())
    const merged = executeOps(editor, [
      { op: 'mergeTableCells', target, range: { row: 0, col: 0, rowEnd: 1, colEnd: 1 } },
    ])
    expect(merged.ok, merged.error).toBe(true)
    expect(rowTexts(editor)).toEqual([['A1B1A2B2', 'C1'], ['C2'], ['A3', 'B3', 'C3']])

    const res = executeOps(editor, [{ op: 'deleteTableColumn', target, col: 2 }])
    expect(res.ok).toBe(false)
    expect(res.error).toContain('would leave row 1 without cells')
    expect(rowTexts(editor)).toEqual([['A1B1A2B2', 'C1'], ['C2'], ['A3', 'B3', 'C3']])

    const ok = executeOps(editor, [{ op: 'deleteTableColumn', target, col: 0, count: 2 }])
    expect(ok.ok, ok.error).toBe(true)
    expect(rowTexts(editor)).toEqual([['C1'], ['C2'], ['C3']])
  })

  it('merges a rectangle, refuses a range that cuts a merged cell, and splits again', () => {
    const editor = createEditor(doc())
    const merge = executeOps(editor, [
      { op: 'mergeTableCells', target, range: { row: 1, col: 0, rowEnd: 2, colEnd: 1 } },
    ])
    expect(merge.ok, merge.error).toBe(true)
    const merged = tableNode(editor).child(1).child(0)
    expect(merged.attrs).toMatchObject({ colspan: 2, rowspan: 2, colwidth: [100, 100] })
    expect(merged.textContent).toBe('A2B2A3B3')
    expect(merged.childCount).toBe(4)
    expect(describeTable(tableNode(editor))).toEqual({
      rows: 3,
      cols: 3,
      merged: [{ row: 1, col: 0, rowSpan: 2, colSpan: 2 }],
    })
    const model = pmTableToModel(editor.getJSON().content![1] as never)
    expect(model.rows[1][0]).toMatchObject({ colSpan: 2, vMerge: 'restart' })
    expect(model.rows[2][0]).toMatchObject({ colSpan: 2, vMerge: 'continue' })
    expect(model.rows[2]).toHaveLength(2)

    const cut = executeOps(editor, [
      { op: 'mergeTableCells', target, range: { row: 0, col: 0, rowEnd: 1, colEnd: 0 } },
    ])
    expect(cut.ok).toBe(false)
    expect(cut.error).toContain('cell at row 1 col 0 spans rows 1-2, cols 0-1')
    expect(tableNode(editor).child(1).child(0).attrs.colspan).toBe(2)

    const single = executeOps(editor, [
      { op: 'mergeTableCells', target, range: { row: 0, col: 0 } },
    ])
    expect(single.error).toContain('single cell')

    const split = executeOps(editor, [{ op: 'splitTableCell', target, cell: { row: 1, col: 0 } }])
    expect(split.ok, split.error).toBe(true)
    expect(rowTexts(editor)).toEqual([
      ['A1', 'B1', 'C1'],
      ['A2B2A3B3', '', 'C2'],
      ['', '', 'C3'],
    ])
    expect(tableNode(editor).child(2).child(0).attrs.colwidth).toEqual([100])
    const notMerged = executeOps(editor, [
      { op: 'splitTableCell', target, cell: { row: 0, col: 0 } },
    ])
    expect(notMerged.error).toContain('is not merged')
  })

  it('formats cells: fill, vAlign, borders and column width with unit suffixes', () => {
    const editor = createEditor(doc())
    const r = executeOps(editor, [
      {
        op: 'setTableCellFormat',
        target,
        range: { row: 0, col: 0, colEnd: 2 },
        fill: '#1F4E79',
        vAlign: 'center',
        borders: { all: { style: 'single', width: 1.5, color: '#FF0000' }, bottom: null },
      },
      { op: 'setTableCellFormat', target, range: { row: 0, col: 1 }, width: '1in' },
    ])
    expect(r.ok, r.error).toBe(true)
    const head = tableNode(editor).child(0).child(0)
    expect(head.attrs.fill).toBe('1F4E79')
    expect(head.attrs.vAlign).toBe('center')
    expect(head.attrs.borders).toEqual({
      top: { style: 'single', szEighths: 12, color: 'FF0000' },
      left: { style: 'single', szEighths: 12, color: 'FF0000' },
      right: { style: 'single', szEighths: 12, color: 'FF0000' },
      bottom: { style: 'none' },
    })
    expect(tableNode(editor).child(2).child(0).attrs.fill).toBeNull()
    tableNode(editor).forEach((row) => expect(row.child(1).attrs.colwidth).toEqual([96]))
    expect(tableNode(editor).attrs.widthPx).toBe(296)

    const clear = executeOps(editor, [
      { op: 'setTableCellFormat', target, fill: null, vAlign: null, borders: null },
    ])
    expect(clear.ok).toBe(true)
    tableNode(editor).forEach((row) =>
      row.forEach((c) => {
        expect(c.attrs.fill).toBeNull()
        expect(c.attrs.vAlign).toBeNull()
        expect(c.attrs.borders).toEqual({
          top: { style: 'none' },
          left: { style: 'none' },
          bottom: { style: 'none' },
          right: { style: 'none' },
        })
      }),
    )
    expect(parseTableLength('2.5cm')).toBe(1417.5)
    expect(parseTableLength(720)).toBe(720)
    expect(parseTableLength('12em')).toBeUndefined()
  })

  it('refuses a merge that would leave a row without cells, keeps header cells on append, width null is invalid', () => {
    const editor = createEditor(doc())
    const wide = executeOps(editor, [
      { op: 'mergeTableCells', target, range: { row: 0, col: 0, rowEnd: 1, colEnd: 2 } },
    ])
    expect(wide.ok).toBe(false)
    expect(wide.error).toContain('would leave row 1 without cells')
    expect(rowTexts(editor)).toHaveLength(3)

    const header = createEditor([
      para('x'),
      {
        ...table(),
        content: [
          {
            type: 'docTableRow',
            content: [
              { ...cell('H1'), type: 'docTableHeader' },
              { ...cell('H2'), type: 'docTableHeader' },
            ],
          },
          { type: 'docTableRow', content: [cell('a'), cell('b')] },
        ],
      },
    ])
    const appended = executeOps(header, [
      { op: 'insertTableColumn', target, col: 1 },
      { op: 'insertTableRow', target, row: 1, count: 2 },
    ])
    expect(appended.ok, appended.error).toBe(true)
    expect(tableNode(header).child(0).child(2).type.name).toBe('docTableHeader')
    expect(tableNode(header).child(1).child(2).type.name).toBe('docTableCell')
    expect(tableNode(header).childCount).toBe(4)

    const nullWidth = executeOps(editor, [{ op: 'setTableCellFormat', target, width: null }])
    expect(nullWidth.ok).toBe(false)
    expect(nullWidth.error).toContain('width must be a positive number of twips')
  })

  it('rows inserted next to a header row are body rows; a rowspan band grows and keeps the free column', () => {
    const header = createEditor([
      para('x'),
      {
        ...table(),
        content: [
          {
            type: 'docTableRow',
            content: [
              { ...cell('H1'), type: 'docTableHeader' },
              { ...cell('H2'), type: 'docTableHeader' },
            ],
          },
          { type: 'docTableRow', content: [cell('a'), cell('b')] },
        ],
      },
    ])
    const r = executeOps(header, [
      { op: 'insertTableRow', target, row: 0 },
      { op: 'insertTableRow', target, row: 0, position: 'before' },
    ])
    expect(r.ok, r.error).toBe(true)
    const t = tableNode(header)
    expect(t.childCount).toBe(4)
    expect(t.child(0).child(0).type.name).toBe('docTableCell')
    expect(t.child(1).child(0).type.name).toBe('docTableHeader')
    expect(t.child(2).child(0).type.name).toBe('docTableCell')

    // a merge spanning every column is refused upstream, so the widest band leaves one column free
    const editor = createEditor(doc())
    const band = executeOps(editor, [
      { op: 'mergeTableCells', target, range: { row: 1, col: 0, rowEnd: 2, colEnd: 1 } },
      { op: 'insertTableRow', target, row: 1 },
    ])
    expect(band.ok, band.error).toBe(true)
    expect(tableNode(editor).childCount).toBe(4)
    expect(tableNode(editor).child(1).child(0).attrs.rowspan).toBe(3)
    expect(tableNode(editor).child(2).childCount).toBe(1)
    expect(describeTable(tableNode(editor)).merged).toEqual([
      { row: 1, col: 0, rowSpan: 3, colSpan: 2 },
    ])
  })

  it('a rowspan next to an inserted row grows once per inserted row, after earlier batch steps', () => {
    const editor = createEditor(doc())
    const r = executeOps(editor, [
      { op: 'insertTableColumn', target, col: 0, position: 'before' },
      { op: 'mergeTableCells', target, range: { row: 0, col: 3, rowEnd: 2 } },
      { op: 'insertTableRow', target, row: 0, count: 2 },
    ])
    expect(r.ok, r.error).toBe(true)
    const t = tableNode(editor)
    expect(t.childCount).toBe(5)
    expect(t.child(0).child(3).attrs.rowspan).toBe(5)
    expect(t.child(1).childCount).toBe(3)
    expect(t.child(2).childCount).toBe(3)
    expect(describeTable(t).merged).toEqual([{ row: 0, col: 3, rowSpan: 5, colSpan: 1 }])
  })

  it('validates shapes up front without touching the document', () => {
    const editor = createEditor(doc())
    const bad = executeOps(editor, [
      { op: 'insertTableRow', target, row: -1 },
      { op: 'mergeTableCells', target, range: { row: 1, col: 1, rowEnd: 0 } },
      { op: 'setTableCellFormat', target, vAlign: 'middle' },
      { op: 'setTableCellFormat', target, width: '3em' },
      { op: 'setTableCellFormat', target, borders: { diagonal: {} } },
      { op: 'setTableStyle', target },
    ])
    expect(bad.ok).toBe(false)
    for (const piece of [
      'row must be a non-negative integer',
      'range.rowEnd must be >= row',
      'vAlign must be top / center / bottom',
      'width must be a positive number of twips',
      'unknown borders side "diagonal"',
      'give styleId and/or look',
    ]) {
      expect(bad.error).toContain(piece)
    }
    expect(rowTexts(editor)[0]).toEqual(['A1', 'B1', 'C1'])
  })

  it('needs exactly one native table; nodeType "table" selects them', () => {
    const editor = createEditor([
      para('x'),
      table(),
      {
        type: 'docProtected',
        attrs: { docxIndex: 3, blockType: 'table', label: 'Table', previewText: 'ro' },
      },
      table({ docxIndex: 9 }),
    ])
    const two = executeOps(editor, [
      { op: 'insertTableRow', target: { nodeType: 'table' }, row: 0 },
    ])
    expect(two.ok).toBe(false)
    expect(two.error).toContain('matches 2 tables (blocks 1, 3)')
    const ro = executeOps(editor, [{ op: 'insertTableRow', target: { blockIndexes: [2] }, row: 0 }])
    expect(ro.error).toContain('matches no native table (matched: read-only table)')
    const none = executeOps(editor, [
      { op: 'insertTableRow', target: { blockIndexes: [42] }, row: 0 },
    ])
    expect(none.ok).toBe(true)
    expect(none.results[0]).toMatchObject({ matched: 0, changed: 0 })
    const one = executeOps(editor, [
      {
        op: 'insertTableRow',
        target: { nodeType: 'table', containsText: 'A1', blockIndexes: [3] },
        row: 2,
      },
    ])
    expect(one.ok, one.error).toBe(true)
    expect(editor.state.doc.child(3).childCount).toBe(4)
    expect(editor.state.doc.child(1).childCount).toBe(3)
  })

  it('sets the table style from styles.xml and the look flags', () => {
    const editor = createEditor(doc())
    editor.storage.listNumbering.styles = new Map([
      ['TableGrid', { styleId: 'TableGrid', name: 'Table Grid', type: 'table' as const }],
      ['GridTable4', { styleId: 'GridTable4', name: 'Grid Table 4', type: 'table' as const }],
      ['Heading1', { styleId: 'Heading1', name: 'heading 1', type: 'paragraph' as const }],
    ])
    const missing = executeOps(editor, [{ op: 'setTableStyle', target, styleId: 'Fancy' }])
    expect(missing.ok).toBe(false)
    expect(missing.error).toContain('no table style "Fancy"')
    expect(missing.error).toContain('Available: [TableGrid, GridTable4]')
    const r = executeOps(editor, [
      {
        op: 'setTableStyle',
        target,
        styleId: 'grid table 4',
        look: { bandedRows: false, lastRow: true },
      },
    ])
    expect(r.ok, r.error).toBe(true)
    expect(tableNode(editor).attrs).toMatchObject({
      tblStyleId: 'GridTable4',
      tblLookEdited: true,
      tblLook: {
        firstRow: true,
        lastRow: true,
        firstColumn: true,
        lastColumn: false,
        bandedRows: false,
        bandedColumns: false,
      },
    })
    const model = pmTableToModel(editor.getJSON().content![1] as never)
    expect(model.tblStyleId).toBe('GridTable4')
    expect(model.tableLook).toMatchObject({ bandedRows: false, lastRow: true })
    const cleared = executeOps(editor, [{ op: 'setTableStyle', target, styleId: null }])
    expect(cleared.ok).toBe(true)
    expect(pmTableToModel(editor.getJSON().content![1] as never).tblStyleId).toBe('')
  })

  it('a later op in the batch sees the earlier structural change', () => {
    const editor = createEditor(doc())
    const r = executeOps(editor, [
      { op: 'insertTableColumn', target, col: 0, position: 'before' },
      { op: 'insertTableRow', target, row: 0, position: 'before' },
      { op: 'mergeTableCells', target, range: { row: 0, col: 0, colEnd: 3 } },
      { op: 'setTableCellFormat', target, range: { row: 0, col: 0 }, fill: '#000000' },
    ])
    expect(r.ok, r.error).toBe(true)
    expect(describeTable(tableNode(editor))).toEqual({
      rows: 4,
      cols: 4,
      merged: [{ row: 0, col: 0, rowSpan: 1, colSpan: 4 }],
    })
    expect(tableNode(editor).child(0).child(0).attrs.fill).toBe('000000')
    expect(rowTexts(editor)[1]).toEqual(['', 'A1', 'B1', 'C1'])
  })
})
