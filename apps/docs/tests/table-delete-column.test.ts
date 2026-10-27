import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { deleteTableColumn, selectTablePart } from '../src/renderer/editor/table-ops'

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

/** A=100px, B=120px, C=80px (300 total); colWidthsPct states the same grid */
const table = (): JsonNode => ({
  type: 'docTable',
  attrs: { docxIndex: 7, widthPx: 300, colWidthsPct: [33.33, 40, 26.67] },
  content: ['1', '2', '3'].map((r) => ({
    type: 'docTableRow',
    content: [
      cell(`A${r}`, { colwidth: [100] }),
      cell(`B${r}`, { colwidth: [120] }),
      cell(`C${r}`, { colwidth: [80] }),
    ],
  })),
})

/** first row has a cell merged across A+B, as Word writes spans */
const mergedTable = (): JsonNode => ({
  type: 'docTable',
  attrs: { docxIndex: 7, widthPx: 300, colWidthsPct: [33.33, 40, 26.67] },
  content: [
    {
      type: 'docTableRow',
      content: [cell('X', { colspan: 2, colwidth: [100, 120] }), cell('C1', { colwidth: [80] })],
    },
    {
      type: 'docTableRow',
      content: [
        cell('A2', { colwidth: [100] }),
        cell('B2', { colwidth: [120] }),
        cell('C2', { colwidth: [80] }),
      ],
    },
  ],
})

/** a percentage-based table carries no px widths at all */
const pctTable = (): JsonNode => ({
  type: 'docTable',
  attrs: { docxIndex: 7, colWidthsPct: [20, 30, 50] },
  content: ['1', '2'].map((r) => ({
    type: 'docTableRow',
    content: [
      cell(`A${r}`, { colwidth: null }),
      cell(`B${r}`, { colwidth: null }),
      cell(`C${r}`, { colwidth: null }),
    ],
  })),
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function createEditor(content: JsonNode): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('Intro'), content, para('Outro')] },
  })
  editors.add(editor)
  return editor
}

function caretInto(editor: Editor, label: string): void {
  let pos = 0
  editor.state.doc.descendants((node, p) => {
    if (!pos && node.type.name === 'docTableCell' && node.textContent === label) pos = p
    return !pos
  })
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 2)),
  )
}

const deleteColumnOf = (editor: Editor, label: string): void => {
  caretInto(editor, label)
  expect(selectTablePart('column')(editor.state, editor.view.dispatch)).toBe(true)
  expect(deleteTableColumn(editor.state, editor.view.dispatch)).toBe(true)
}

const tableNode = (editor: Editor) => editor.state.doc.child(1)

const rowTexts = (editor: Editor): string[][] => {
  const out: string[][] = []
  tableNode(editor).forEach((row) => {
    const cells: string[] = []
    row.forEach((c) => cells.push(c.textContent))
    out.push(cells)
  })
  return out
}

const firstRowWidths = (editor: Editor): Array<number[] | null> => {
  const out: Array<number[] | null> = []
  tableNode(editor).firstChild?.forEach((c) => out.push(c.attrs.colwidth as number[] | null))
  return out
}

const pctRounded = (editor: Editor): number[] =>
  (tableNode(editor).attrs.colWidthsPct as number[]).map((p) => Math.round(p))

describe('table delete column keeps the width grid', () => {
  it('deleting the leftmost column keeps the surviving px widths and shrinks the table', () => {
    const editor = createEditor(table())
    deleteColumnOf(editor, 'A2')
    expect(rowTexts(editor)).toEqual([
      ['B1', 'C1'],
      ['B2', 'C2'],
      ['B3', 'C3'],
    ])
    // the surviving cells keep their own px widths (Word behavior)
    expect(firstRowWidths(editor)).toEqual([[120], [80]])
    // the table shrinks by the deleted column; the pct grid loses its share
    expect(tableNode(editor).attrs.widthPx).toBe(200)
    expect(pctRounded(editor)).toEqual([60, 40])
  })

  it('deleting a middle column splices that column out of the grid', () => {
    const editor = createEditor(table())
    deleteColumnOf(editor, 'B2')
    expect(rowTexts(editor)).toEqual([
      ['A1', 'C1'],
      ['A2', 'C2'],
      ['A3', 'C3'],
    ])
    expect(firstRowWidths(editor)).toEqual([[100], [80]])
    expect(tableNode(editor).attrs.widthPx).toBe(180)
    // [33.33, 26.67] renormalized over the survivors
    expect(pctRounded(editor)).toEqual([56, 44])
  })

  it('a merged cell spanning the deleted column shrinks to the surviving grid', () => {
    const editor = createEditor(mergedTable())
    // caret delete (no column selection): with the merge in the first row a
    // column selection would span both grid columns, which deletes wider
    caretInto(editor, 'B2')
    expect(deleteTableColumn(editor.state, editor.view.dispatch)).toBe(true)
    expect(rowTexts(editor)).toEqual([
      ['X', 'C1'],
      ['A2', 'C2'],
    ])
    const x = tableNode(editor).firstChild?.firstChild
    expect(Number(x?.attrs.colspan)).toBe(1)
    expect(x?.attrs.colwidth).toEqual([100])
    expect(tableNode(editor).attrs.widthPx).toBe(180)
    expect(pctRounded(editor)).toEqual([56, 44])
  })

  it('a percentage grid renormalizes onto the survivors', () => {
    const editor = createEditor(pctTable())
    deleteColumnOf(editor, 'A2')
    expect(rowTexts(editor)).toEqual([
      ['B1', 'C1'],
      ['B2', 'C2'],
    ])
    // no px widths existed, none appear; the shares fill the same widthPct
    expect(firstRowWidths(editor)).toEqual([null, null])
    expect(tableNode(editor).attrs.widthPx).toBeNull()
    // [30, 50] renormalized (37.5, 62.5)
    expect(pctRounded(editor)).toEqual([38, 63])
  })
})
