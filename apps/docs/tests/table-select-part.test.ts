import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import {
  enterSelectedTable,
  inTableOrSelected,
  selectTablePart,
  tableCellsSelection,
} from '../src/renderer/editor/table-ops'
import { repeatHeaderState } from '../src/renderer/editor/table-properties'

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
const cell = (t: string): JsonNode => ({
  type: 'docTableCell',
  attrs: { colspan: 1, rowspan: 1, colwidth: [100] },
  content: [para(t)],
})
const table = (): JsonNode => ({
  type: 'docTable',
  attrs: { docxIndex: 7, widthPx: 300 },
  content: ['1', '2'].map((r) => ({
    type: 'docTableRow',
    content: ['A', 'B', 'C'].map((c) => cell(`${c}${r}`)),
  })),
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('Intro'), table(), para('Outro')] },
  })
  editors.add(editor)
  // caret in B1
  let pos = 0
  editor.state.doc.descendants((node, p) => {
    if (!pos && node.type.name === 'docTableCell' && node.textContent === 'B1') pos = p
    return !pos
  })
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 2)),
  )
  return editor
}

const run = (editor: Editor, kind: Parameters<typeof selectTablePart>[0]) =>
  selectTablePart(kind)(editor.state, editor.view.dispatch)

describe('Table Layout ▸ Select ▾', () => {
  it('Select Table keeps the table contextual state alive', () => {
    const editor = createEditor()
    expect(run(editor, 'table')).toBe(true)
    const sel = editor.state.selection
    expect(sel).toBeInstanceOf(NodeSelection)
    expect((sel as NodeSelection).node.type.name).toBe('docTable')
    expect(inTableOrSelected(editor.state)).toBe(true)
    const fs = computeFormatState(editor)
    expect(fs.inTable).toBe(true)
    expect(fs.canMergeCells).toBe(true)
    expect(fs.cellKey).not.toBeNull()
    expect(repeatHeaderState(editor.state).enabled).toBe(false)
    const cells = tableCellsSelection(editor.state)!
    expect(repeatHeaderState(editor.state.apply(editor.state.tr.setSelection(cells))).enabled).toBe(
      true,
    )
  })

  it('ribbon commands first turn the selected table into a selection of every cell', () => {
    const editor = createEditor()
    run(editor, 'table')
    expect(run(editor, 'row')).toBe(false)
    expect(enterSelectedTable(editor.state, editor.view.dispatch)).toBe(true)
    let sel = editor.state.selection as CellSelection
    expect(sel).toBeInstanceOf(CellSelection)
    expect(sel.$anchorCell.nodeAfter?.textContent).toBe('A1')
    expect(sel.$headCell.nodeAfter?.textContent).toBe('C2')
    expect(run(editor, 'row')).toBe(true)
    sel = editor.state.selection as CellSelection
    expect(sel.$anchorCell.nodeAfter?.textContent).toBe('A1')
    expect(sel.$headCell.nodeAfter?.textContent).toBe('C1')
  })

  it('enterSelectedTable leaves an in-table caret and a plain paragraph alone', () => {
    const editor = createEditor()
    const before = editor.state.selection
    expect(enterSelectedTable(editor.state, editor.view.dispatch)).toBe(true)
    expect(editor.state.selection.eq(before)).toBe(true)
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)))
    expect(enterSelectedTable(editor.state, editor.view.dispatch)).toBe(false)
  })

  it('row / column / cell select from the caret cell', () => {
    const editor = createEditor()
    run(editor, 'column')
    let sel = editor.state.selection as CellSelection
    expect(sel.$anchorCell.nodeAfter?.textContent).toBe('B1')
    expect(sel.$headCell.nodeAfter?.textContent).toBe('B2')
    run(editor, 'cell')
    sel = editor.state.selection as CellSelection
    expect(sel.$anchorCell.pos).toBe(sel.$headCell.pos)
    expect(sel.$anchorCell.nodeAfter?.textContent).toBe('B1')
  })

  it('outside a table nothing is selected', () => {
    const editor = createEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)))
    expect(run(editor, 'table')).toBe(false)
    expect(inTableOrSelected(editor.state)).toBe(false)
  })
})
