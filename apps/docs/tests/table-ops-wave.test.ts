import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection, type Command } from '@tiptap/pm/state'
import { CellSelection, TableMap, mergeCells } from '@tiptap/pm/tables'
import type { Node as PmNode } from '@tiptap/pm/model'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  type PmNode as JsonNode,
} from '../src/renderer/editor/convert'
import {
  canShiftCells,
  deleteCells,
  editableOrToast,
  insertRowsOrColumns,
  selectTablePart,
  setCellAlignment,
  setCellProps,
  setCellTextDirection,
  setRowProps,
  splitCellsInto,
  splitGridCell,
  splitTableAtSelection,
  distributeRowsEvenly,
} from '../src/renderer/editor/table-ops'
import { distributeSelectedColumns } from '../src/renderer/editor/table-sizing'
import { updateSelectedTableAttrs } from '../src/renderer/editor/table-properties'
import { overlayLiveRowFlags } from '../src/renderer/pagination-lines'
import { setToastEmitter } from '../src/renderer/components/toast-bus'
import { rowPatchFor } from '../src/renderer/components/TablePropertiesDialog'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
  vi.restoreAllMocks()
})

const cellXml = (text: string, tcPr = '') =>
  `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`

/** rows × cols grid, cells named A1..; optional per-row trPr and per-cell tcPr */
function tableXml(
  rows: number,
  cols: number,
  opts: { trPr?: Record<number, string>; tcPr?: Record<string, string>; tblPr?: string } = {},
): string {
  const grid = Array.from({ length: cols }, () => '<w:gridCol w:w="2000"/>').join('')
  const body = Array.from({ length: rows }, (_, r) => {
    const cells = Array.from({ length: cols }, (_, c) => {
      const name = `${String.fromCharCode(65 + c)}${r + 1}`
      return cellXml(name, opts.tcPr?.[name] ?? '')
    }).join('')
    return `<w:tr>${opts.trPr?.[r] ?? ''}${cells}</w:tr>`
  }).join('')
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>${opts.tblPr ?? ''}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`
}

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml: '<w:p/>' + bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, parsed }
}

function tableAt(editor: Editor, index = 0): { node: PmNode; pos: number } {
  let found: { node: PmNode; pos: number } | null = null
  let seen = 0
  editor.state.doc.forEach((node, offset) => {
    if (found || node.type.name !== 'docTable') return
    if (seen++ === index) found = { node, pos: offset }
  })
  if (!found) throw new Error('no table')
  return found
}

function cellPos(editor: Editor, row: number, col: number, index = 0): number {
  const { node, pos } = tableAt(editor, index)
  const map = TableMap.get(node)
  return pos + 1 + map.map[row * map.width + col]
}

function caretIn(editor: Editor, row: number, col: number) {
  const pos = cellPos(editor, row, col)
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(pos + 1))),
  )
}

function selectCells(editor: Editor, from: [number, number], to: [number, number]) {
  const doc = editor.state.doc
  editor.view.dispatch(
    editor.state.tr.setSelection(
      CellSelection.create(doc, cellPos(editor, from[0], from[1]), cellPos(editor, to[0], to[1])),
    ),
  )
}

const run = (editor: Editor, command: Command) => command(editor.state, editor.view.dispatch)

function grid(editor: Editor, index = 0): string[][] {
  const { node } = tableAt(editor, index)
  const map = TableMap.get(node)
  const out: string[][] = []
  for (let r = 0; r < map.height; r++) {
    out.push(
      Array.from(
        { length: map.width },
        (_, c) => node.nodeAt(map.map[r * map.width + c])!.textContent,
      ),
    )
  }
  return out
}

function savedTableXml(
  editor: Editor,
  parsed: Awaited<ReturnType<typeof parseDocx>>,
  index = 0,
): string {
  const plan = pmDocToSavePlan(editor.getJSON() as JsonNode, parsed.blocks)
  const tables = plan.saveBlocks.filter((b) => b.kind === 'xml' && b.xml.includes('<w:tbl>'))
  const block = tables[index]
  if (!block || block.kind !== 'xml') throw new Error('table not regenerated')
  return block.xml
}

describe('splitGridCell', () => {
  it('splits a plain cell into columns: the grid widens inside the span, neighbours grow their colspan', async () => {
    const { editor } = await open(tableXml(2, 2))
    const { node } = tableAt(editor)
    const result = splitGridCell(node, { row: 0, col: 0 }, 1, 3)!
    const map = TableMap.get(result.table)
    expect([map.width, map.height]).toEqual([4, 2])
    const rowBelow = result.table.child(1)
    expect(rowBelow.child(0).attrs.colspan).toBe(3)
    expect(rowBelow.child(1).attrs.colspan).toBe(1)
    const rowTop = result.table.child(0)
    expect(rowTop.childCount).toBe(4)
    expect(rowTop.child(0).textContent).toBe('A1')
    expect(rowTop.child(1).textContent).toBe('')
    // widths: the split cell's 2000-twip column becomes three thirds, B1 keeps its own
    const widths = (i: number) => rowTop.child(i).attrs.colwidth as number[]
    expect(widths(0)[0]).toBeCloseTo(widths(1)[0], 5)
    expect(widths(0)[0] * 3).toBeCloseTo(
      (rowBelow.child(0).attrs.colwidth as number[]).reduce((a, b) => a + b, 0),
      3,
    )
    expect(result.table.attrs.colWidthsPct).toHaveLength(4)
  })

  it('splits into rows: new rows go under the cell and the neighbour spans them', async () => {
    const { editor } = await open(tableXml(2, 2))
    const { node } = tableAt(editor)
    const result = splitGridCell(node, { row: 0, col: 0 }, 3, 1)!
    const map = TableMap.get(result.table)
    expect([map.width, map.height]).toEqual([2, 4])
    expect(result.table.child(0).child(1).attrs.rowspan).toBe(3)
    expect(result.table.child(0).child(1).textContent).toBe('B1')
    expect(result.table.child(1).childCount).toBe(1)
    expect(result.table.child(3).child(0).textContent).toBe('A2')
  })

  it('splits a merged cell back along its own span without adding columns', async () => {
    const { editor } = await open(tableXml(2, 2))
    selectCells(editor, [0, 0], [0, 1])
    run(editor, mergeCells)
    const { node } = tableAt(editor)
    expect(node.child(0).child(0).attrs.colspan).toBe(2)
    const result = splitGridCell(node, { row: 0, col: 0 }, 1, 2)!
    const map = TableMap.get(result.table)
    expect([map.width, map.height]).toEqual([2, 2])
    expect(result.table.child(0).childCount).toBe(2)
    expect(result.table.child(0).child(0).attrs.colspan).toBe(1)
  })

  it('refuses a 1×1 split', async () => {
    const { editor } = await open(tableXml(2, 2))
    expect(splitGridCell(tableAt(editor).node, { row: 0, col: 0 }, 1, 1)).toBeNull()
  })
})

describe('splitCellsInto command', () => {
  it('merges a multi-cell selection first when asked, then splits the result', async () => {
    const { editor, parsed } = await open(tableXml(2, 2))
    selectCells(editor, [0, 0], [0, 1])
    expect(run(editor, splitCellsInto(2, 3, true))).toBe(true)
    expect(grid(editor).length).toBe(3)
    expect(grid(editor)[0].length).toBe(3)
    expect(grid(editor)[0]).toEqual(['A1B1', '', ''])
    // the two merged grid columns became three; the row below keeps A2 over the first two
    expect(tableAt(editor).node.child(2).child(0).attrs.colspan).toBe(2)
    const xml = savedTableXml(editor, parsed)
    expect(xml.match(/<w:tr>/g)).toHaveLength(3)
    expect(xml).toContain('<w:gridSpan w:val="2"/>')
  })

  it('splits a merged cell under a plain caret without trying to merge again', async () => {
    const { editor } = await open(tableXml(2, 3))
    selectCells(editor, [0, 0], [0, 1])
    run(editor, mergeCells)
    caretIn(editor, 0, 0)
    expect(run(editor, splitCellsInto(2, 3, true))).toBe(true)
    expect(grid(editor).length).toBe(3)
    expect(grid(editor)[0]).toEqual(['A1B1', '', '', 'C1'])
  })

  it('writes the cantSplit the new split rows inherit', async () => {
    const { editor, parsed } = await open(
      tableXml(1, 2, { trPr: { 0: '<w:trPr><w:cantSplit/></w:trPr>' } }),
    )
    caretIn(editor, 0, 0)
    run(editor, splitCellsInto(3, 1, false))
    const xml = savedTableXml(editor, parsed)
    expect(xml.match(/<w:tr>/g)).toHaveLength(3)
    expect(xml.match(/<w:cantSplit\/>/g)).toHaveLength(3)
  })

  it('leaves the caret in the first piece', async () => {
    const { editor } = await open(tableXml(1, 1))
    caretIn(editor, 0, 0)
    run(editor, splitCellsInto(1, 2, false))
    expect(editor.state.selection.$from.parent.textContent).toBe('A1')
  })
})

describe('splitTableAtSelection', () => {
  it('starts a new table at the caret row with an empty paragraph between', async () => {
    const { editor, parsed } = await open(tableXml(3, 2))
    caretIn(editor, 1, 0)
    expect(run(editor, splitTableAtSelection())).toBe(true)
    const kinds: string[] = []
    editor.state.doc.forEach((n) => kinds.push(n.type.name))
    expect(kinds).toEqual(['docParagraph', 'docTable', 'docParagraph', 'docTable'])
    expect(grid(editor, 0)).toEqual([['A1', 'B1']])
    expect(grid(editor, 1)).toEqual([
      ['A2', 'B2'],
      ['A3', 'B3'],
    ])
    expect(editor.state.selection.$from.parent.textContent).toBe('A2')
    expect(savedTableXml(editor, parsed, 1).match(/<w:tr>/g)).toHaveLength(2)
  })

  it('leaves the alt text on the upper table only', async () => {
    const { editor, parsed } = await open(
      tableXml(2, 1, { tblPr: '<w:tblCaption w:val="Totals"/><w:tblDescription w:val="Yearly"/>' }),
    )
    caretIn(editor, 1, 0)
    run(editor, splitTableAtSelection())
    expect(tableAt(editor, 1).node.attrs.tblCaption).toBeNull()
    const lower = savedTableXml(editor, parsed, 1)
    expect(lower).not.toContain('tblCaption')
    expect(lower).not.toContain('tblDescription')
    expect(lower).toContain('<w:tblStyle w:val="TableGrid"/>')
    expect(savedTableXml(editor, parsed, 0)).toContain('<w:tblCaption w:val="Totals"/>')
  })

  it('is refused on the first row and cuts a crossing vertical merge into an empty cell', async () => {
    const { editor } = await open(tableXml(3, 2))
    caretIn(editor, 0, 0)
    expect(run(editor, splitTableAtSelection())).toBe(false)
    selectCells(editor, [0, 0], [2, 0])
    run(editor, mergeCells)
    caretIn(editor, 1, 1)
    expect(run(editor, splitTableAtSelection())).toBe(true)
    const upper = tableAt(editor, 0).node
    const lower = tableAt(editor, 1).node
    expect(upper.child(0).child(0).attrs.rowspan).toBe(1)
    expect(lower.child(0).child(0).attrs.rowspan).toBe(2)
    expect(lower.child(0).child(0).textContent).toBe('')
  })
})

describe('insertRowsOrColumns', () => {
  it('inserts N rows below and N columns to the right with a refitted grid', async () => {
    const { editor } = await open(tableXml(2, 2))
    caretIn(editor, 0, 0)
    run(editor, insertRowsOrColumns(3, 'below', 600))
    expect(grid(editor).length).toBe(5)
    expect(grid(editor)[1]).toEqual(['', ''])
    run(editor, insertRowsOrColumns(2, 'right', 600))
    const { node } = tableAt(editor)
    const map = TableMap.get(node)
    expect(map.width).toBe(4)
    node.child(0).forEach((cell) => expect((cell.attrs.colwidth as number[])[0]).toBeGreaterThan(0))
    const total = (node.attrs.colWidthsPct as number[]).reduce((a, b) => a + b, 0)
    expect(total).toBeCloseTo(100, 3)
  })
})

describe('inserted rows', () => {
  it('copy the neighbour row: cell fill and the trPr height / cantSplit', async () => {
    const { editor, parsed } = await open(
      tableXml(2, 2, {
        trPr: { 0: '<w:trPr><w:cantSplit/><w:trHeight w:val="500"/></w:trPr>' },
        tcPr: { A1: '<w:shd w:val="clear" w:color="auto" w:fill="FFCC00"/>' },
      }),
    )
    caretIn(editor, 0, 0)
    run(editor, insertRowsOrColumns(2, 'below', 600))
    const table = tableAt(editor).node
    expect(table.childCount).toBe(4)
    for (const i of [1, 2]) {
      expect(table.child(i).child(0).attrs.fill).toBe('FFCC00')
      expect(table.child(i).attrs).toMatchObject({ heightTwips: 500, cantSplit: true })
    }
    const xml = savedTableXml(editor, parsed)
    expect(xml.match(/<w:cantSplit\/>/g)).toHaveLength(3)
    expect(xml.match(/w:fill="FFCC00"/g)).toHaveLength(3)
  })
})

describe('split every selected cell', () => {
  it('splits each cell of the selection when the merge box is off', async () => {
    const { editor } = await open(tableXml(2, 2))
    selectCells(editor, [0, 0], [0, 1])
    expect(run(editor, splitCellsInto(1, 2, false))).toBe(true)
    expect(grid(editor)[0]).toEqual(['A1', '', 'B1', ''])
    expect(tableAt(editor).node.child(1).child(0).attrs.colspan).toBe(2)
    expect(tableAt(editor).node.child(1).child(1).attrs.colspan).toBe(2)
  })
})

describe('row dialog patch', () => {
  it('writes only touched fields and never stamps the header flag from an untouched checkbox', () => {
    const row = {
      specifyHeight: true,
      heightCm: 1,
      rule: 'exact' as const,
      allowBreak: true,
      repeatHeader: true,
    }
    expect(rowPatchFor(row, new Set(['heightCm']), true)).toEqual({
      heightTwips: 567,
      heightRule: 'exact',
    })
    expect(rowPatchFor(row, new Set(['repeatHeader']), false)).toEqual({})
    expect(rowPatchFor(row, new Set(['repeatHeader', 'allowBreak']), true)).toEqual({
      cantSplit: false,
      cantSplitEdited: true,
      repeatHeader: true,
      repeatHeaderEdited: true,
    })
  })
})

describe('inserted columns', () => {
  it('carry the neighbour cell direction into the saved file', async () => {
    const { editor, parsed } = await open(
      tableXml(1, 2, { tcPr: { B1: '<w:textDirection w:val="tbRl"/>' } }),
    )
    caretIn(editor, 0, 1)
    run(editor, insertRowsOrColumns(1, 'right', 600))
    const xml = savedTableXml(editor, parsed)
    expect(xml.match(/<w:textDirection w:val="tbRl"\/>/g)).toHaveLength(2)
  })
})

describe('read-only guard', () => {
  it('refuses with a toast when the editor is not editable', async () => {
    const toasts: string[] = []
    setToastEmitter((toast) => toasts.push(toast.kind))
    expect(editableOrToast({ isEditable: true })).toBe(true)
    expect(editableOrToast({ isEditable: false })).toBe(false)
    expect(toasts).toEqual(['error'])
    setToastEmitter(null)
  })
})

describe('deleteCells', () => {
  it('shift left drops the cell and parks a gridAfter placeholder at the row end', async () => {
    const { editor, parsed } = await open(tableXml(2, 3))
    caretIn(editor, 0, 0)
    expect(canShiftCells(editor.state)).toBe(true)
    expect(run(editor, deleteCells('shiftLeft'))).toBe(true)
    expect(grid(editor)[0]).toEqual(['B1', 'C1', ''])
    expect(tableAt(editor).node.child(0).child(2).attrs.gridGap).toBe(true)
    const xml = savedTableXml(editor, parsed)
    expect(xml).toContain('<w:gridAfter w:val="1"/>')
    expect(xml.match(/<w:tr>/)?.length).toBe(1)
  })

  it('shift left on a whole row drops the row; on the whole table the table becomes a paragraph', async () => {
    const { editor, parsed } = await open(tableXml(3, 2))
    selectCells(editor, [1, 0], [1, 1])
    expect(run(editor, deleteCells('shiftLeft'))).toBe(true)
    expect(grid(editor)).toEqual([
      ['A1', 'B1'],
      ['A3', 'B3'],
    ])
    const xml = savedTableXml(editor, parsed)
    for (const tr of xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g)!) expect(tr).toContain('<w:tc>')
    expect(xml).not.toContain('gridAfter')
    selectCells(editor, [0, 0], [1, 1])
    expect(run(editor, deleteCells('shiftLeft'))).toBe(true)
    const kinds: string[] = []
    editor.state.doc.forEach((n) => kinds.push(n.type.name))
    expect(kinds).toEqual(['docParagraph', 'docParagraph'])
  })

  it('keeps the grid edge tags ahead of an existing trPr (cantSplit + height) after a shift left', async () => {
    const { editor, parsed } = await open(
      tableXml(2, 3, { trPr: { 0: '<w:trPr><w:cantSplit/><w:trHeight w:val="600"/></w:trPr>' } }),
    )
    caretIn(editor, 0, 0)
    run(editor, deleteCells('shiftLeft'))
    const xml = savedTableXml(editor, parsed)
    const first = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/)![0]
    expect(first).toMatch(
      /^<w:tr><w:trPr><w:gridAfter w:val="1"\/><w:wAfter w:w="\d+" w:type="dxa"\/><w:cantSplit\/><w:trHeight w:val="600"/,
    )
    expect(first.match(/<w:tc>/g)).toHaveLength(2)
  })

  it('shift up moves the column contents up and empties the bottom cell', async () => {
    const { editor } = await open(tableXml(3, 2))
    caretIn(editor, 0, 1)
    run(editor, deleteCells('shiftUp'))
    expect(grid(editor).map((r) => r[1])).toEqual(['B2', 'B3', ''])
    expect(grid(editor).map((r) => r[0])).toEqual(['A1', 'A2', 'A3'])
  })

  it('refuses to shift in a table with merged cells and still deletes whole rows', async () => {
    const { editor } = await open(tableXml(3, 2))
    selectCells(editor, [0, 0], [0, 1])
    run(editor, mergeCells)
    caretIn(editor, 1, 0)
    expect(canShiftCells(editor.state)).toBe(false)
    expect(run(editor, deleteCells('shiftLeft'))).toBe(false)
    expect(run(editor, deleteCells('row'))).toBe(true)
    expect(grid(editor).length).toBe(2)
  })
})

describe('cell alignment and text direction', () => {
  it('writes vAlign on the cell and jc on its paragraphs', async () => {
    const { editor, parsed } = await open(tableXml(1, 2))
    caretIn(editor, 0, 1)
    run(editor, setCellAlignment('bottom', 'right'))
    const cell = tableAt(editor).node.child(0).child(1)
    expect(cell.attrs.vAlign).toBe('bottom')
    expect(cell.firstChild!.attrs.align).toBe('right')
    const xml = savedTableXml(editor, parsed)
    expect(xml).toContain('<w:vAlign w:val="bottom"/>')
    expect(xml).toContain('<w:jc w:val="right"/>')
    run(editor, setCellAlignment('top', 'left'))
    expect(tableAt(editor).node.child(0).child(1).attrs.vAlign).toBeNull()
  })

  it('saves a text direction and removes a parsed one when set back to horizontal', async () => {
    const { editor, parsed } = await open(
      tableXml(1, 2, { tcPr: { B1: '<w:textDirection w:val="tbRl"/>' } }),
    )
    expect(tableAt(editor).node.child(0).child(1).attrs.textDirection).toBe('tbRl')
    caretIn(editor, 0, 0)
    run(editor, setCellTextDirection('btLr'))
    let xml = savedTableXml(editor, parsed)
    expect(xml).toContain('<w:textDirection w:val="btLr"/>')
    expect(xml).toContain('<w:textDirection w:val="tbRl"/>')
    caretIn(editor, 0, 1)
    run(editor, setCellTextDirection('lrTb'))
    xml = savedTableXml(editor, parsed)
    expect(xml).not.toContain('tbRl')
    expect(xml).toContain('<w:textDirection w:val="btLr"/>')
    expect(editor.view.dom.querySelectorAll('td .cell-vert')).toHaveLength(1)
  })
})

describe('row properties', () => {
  it('publishes cantSplit toggles on the tr so live pagination picks them up over cached samples', async () => {
    const { editor } = await open(tableXml(2, 1))
    caretIn(editor, 0, 0)
    run(editor, setRowProps({ cantSplit: true, cantSplitEdited: true }))
    const table = editor.view.dom.querySelector('table') as HTMLElement
    expect(table.querySelector('tr')!.getAttribute('data-cant-split')).toBe('1')
    const rows = [{ height: 10 }, { height: 10 }] as Parameters<typeof overlayLiveRowFlags>[1]
    overlayLiveRowFlags(table, rows)
    expect(rows.map((r) => r.cantSplit)).toEqual([true, false])
    run(editor, setRowProps({ cantSplit: false, cantSplitEdited: true }))
    overlayLiveRowFlags(editor.view.dom.querySelector('table') as HTMLElement, rows)
    expect(rows[0].cantSplit).toBe(false)
  })

  it('parses cantSplit and writes cantSplit / trHeight from the dialog values', async () => {
    const { editor, parsed } = await open(
      tableXml(2, 1, { trPr: { 0: '<w:trPr><w:cantSplit/></w:trPr>' } }),
    )
    const rows = tableAt(editor).node
    expect(rows.child(0).attrs.cantSplit).toBe(true)
    expect(rows.child(1).attrs.cantSplit).toBe(false)
    expect(editor.view.dom.querySelector('tr')!.getAttribute('data-cant-split')).toBe('1')
    caretIn(editor, 1, 0)
    run(
      editor,
      setRowProps({
        cantSplit: true,
        cantSplitEdited: true,
        heightTwips: 600,
        heightRule: 'exact',
      }),
    )
    caretIn(editor, 0, 0)
    run(editor, setRowProps({ cantSplit: false, cantSplitEdited: true }))
    const xml = savedTableXml(editor, parsed)
    const trs = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g)!
    expect(trs[0]).not.toContain('cantSplit')
    expect(trs[1]).toContain(
      '<w:trPr><w:cantSplit/><w:trHeight w:val="600" w:hRule="exact"/></w:trPr>',
    )
  })
})

describe('cell properties', () => {
  it('writes tcMar and noWrap, parses noWrap back, and clears margins to the table default', async () => {
    const { editor, parsed } = await open(tableXml(1, 2, { tcPr: { B1: '<w:noWrap/>' } }))
    expect(tableAt(editor).node.child(0).child(1).attrs.noWrap).toBe(true)
    caretIn(editor, 0, 0)
    run(editor, setCellProps({ cellMar: { top: 100, left: 200 }, noWrap: true }))
    let xml = savedTableXml(editor, parsed)
    expect(xml).toContain(
      '<w:noWrap/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="200" w:type="dxa"/></w:tcMar>',
    )
    expect(editor.view.dom.querySelector('td')!.style.whiteSpace).toBe('nowrap')
    run(editor, setCellProps({ cellMar: null, noWrap: null }))
    xml = savedTableXml(editor, parsed)
    const first = xml.match(/<w:tc>[\s\S]*?<\/w:tc>/)![0]
    expect(first).not.toContain('tcMar')
    expect(first).not.toContain('noWrap')
  })
})

describe('alt text', () => {
  it('round-trips tblCaption / tblDescription and removes them when cleared', async () => {
    const { editor, parsed } = await open(
      tableXml(1, 1, { tblPr: '<w:tblLook w:val="04A0"/><w:tblCaption w:val="Sales"/>' }),
    )
    expect(tableAt(editor).node.attrs.tblCaption).toBe('Sales')
    caretIn(editor, 0, 0)
    run(
      editor,
      updateSelectedTableAttrs({
        tblCaption: 'Sales 2026',
        tblDescription: 'Quarterly totals',
        tblAltEdited: true,
      }),
    )
    let xml = savedTableXml(editor, parsed)
    expect(xml).toContain(
      '<w:tblLook w:val="04A0"/><w:tblCaption w:val="Sales 2026"/><w:tblDescription w:val="Quarterly totals"/></w:tblPr>',
    )
    run(
      editor,
      updateSelectedTableAttrs({ tblCaption: null, tblDescription: null, tblAltEdited: true }),
    )
    xml = savedTableXml(editor, parsed)
    expect(xml).not.toContain('tblCaption')
    expect(xml).not.toContain('tblDescription')
  })
})

describe('distribute and select', () => {
  it('distributes column widths evenly and row heights from the measured average', async () => {
    const { editor } = await open(tableXml(2, 3))
    const { node, pos } = tableAt(editor)
    let tr = editor.state.tr
    const widths = [[100], [200], [300]]
    // every row: the grid-snap plugin would otherwise pull a lone row back to the column consensus
    node.forEach((row, rowOffset) => {
      row.forEach((cell, offset, i) => {
        tr = tr.setNodeMarkup(pos + 1 + rowOffset + 1 + offset, undefined, {
          ...cell.attrs,
          colwidth: widths[i],
        })
      })
    })
    editor.view.dispatch(tr)
    caretIn(editor, 0, 0)
    run(editor, distributeSelectedColumns(1000))
    const after = tableAt(editor).node.child(0)
    after.forEach((cell) => expect((cell.attrs.colwidth as number[])[0]).toBeCloseTo(200, 3))

    const heights = [20, 40]
    let call = 0
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(
      () => ({ height: heights[call++ % 2] }) as DOMRect,
    )
    expect(distributeRowsEvenly(editor.view)).toBe(true)
    tableAt(editor).node.forEach((row) => {
      expect(row.attrs.heightTwips).toBe(450)
      expect(row.attrs.heightRule).toBe('atLeast')
    })
  })

  it('selects row / column / cell / table from the caret', async () => {
    const { editor } = await open(tableXml(2, 2))
    caretIn(editor, 1, 1)
    run(editor, selectTablePart('row'))
    expect(editor.state.selection).toBeInstanceOf(CellSelection)
    expect((editor.state.selection as CellSelection).ranges).toHaveLength(2)
    run(editor, selectTablePart('cell'))
    expect((editor.state.selection as CellSelection).ranges).toHaveLength(1)
    run(editor, selectTablePart('table'))
    expect(editor.state.selection.constructor.name).toBe('NodeSelection')
  })
})
