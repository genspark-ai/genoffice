import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  TABLE_BORDER_MENU,
  setSelectionBorders,
  sideLine,
  type BorderLine,
} from '../src/renderer/editor/table-borders'

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

const table = (attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docTable',
  attrs: { docxIndex: 7, widthPx: 300, ...attrs },
  content: ['1', '2', '3'].map((r) => ({
    type: 'docTableRow',
    content: ['A', 'B', 'C'].map((c) => cell(`${c}${r}`)),
  })),
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function createEditor(tableAttrs: Record<string, unknown> = {}): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('Intro'), table(tableAttrs), para('Outro')] },
  })
  editors.add(editor)
  return editor
}

const LINE: BorderLine = { style: 'single', szEighths: 12, color: 'FF0000' }

/** positions of the 9 cells as [row][col] */
function cellPositions(editor: Editor): number[][] {
  const out: number[][] = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTableRow') out.push([])
    if (node.type.name === 'docTableCell') out[out.length - 1].push(pos)
    return true
  })
  return out
}

function bordersAt(editor: Editor, row: number, col: number): Record<string, BorderLine> | null {
  const pos = cellPositions(editor)[row][col]
  return editor.state.doc.nodeAt(pos)!.attrs.borders as Record<string, BorderLine> | null
}

function caretIn(editor: Editor, row: number, col: number) {
  const pos = cellPositions(editor)[row][col]
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 2)),
  )
}

function selectCells(editor: Editor, from: [number, number], to: [number, number]) {
  const cells = cellPositions(editor)
  editor.view.dispatch(
    editor.state.tr.setSelection(
      CellSelection.create(editor.state.doc, cells[from[0]][from[1]], cells[to[0]][to[1]]),
    ),
  )
}

function run(editor: Editor, mode: (typeof TABLE_BORDER_MENU)[number]): boolean {
  return setSelectionBorders(mode, LINE)(editor.state, editor.view.dispatch)
}

describe('Borders gallery order', () => {
  it('lists Word entries bottom-first with No Border before the compound modes', () => {
    expect(TABLE_BORDER_MENU).toEqual([
      'bottom',
      'top',
      'left',
      'right',
      'none',
      'all',
      'outer',
      'inner',
      'insideH',
      'insideV',
    ])
  })
})

describe('sideLine mapping', () => {
  const edge = { top: true, bottom: false, left: true, right: false }
  it('single-side modes only touch their own edge of the selection', () => {
    expect(sideLine('bottom', 'bottom', edge, LINE)).toBeUndefined()
    expect(sideLine('top', 'top', edge, LINE)).toBe(LINE)
    expect(sideLine('top', 'bottom', edge, LINE)).toBeUndefined()
  })
  it('inner / inside modes skip the selection edges and split by orientation', () => {
    expect(sideLine('inner', 'top', edge, LINE)).toBeUndefined()
    expect(sideLine('inner', 'bottom', edge, LINE)).toBe(LINE)
    expect(sideLine('insideH', 'bottom', edge, LINE)).toBe(LINE)
    expect(sideLine('insideH', 'right', edge, LINE)).toBeUndefined()
    expect(sideLine('insideV', 'right', edge, LINE)).toBe(LINE)
    expect(sideLine('insideV', 'bottom', edge, LINE)).toBeUndefined()
  })
  it('none clears every side, all paints every side', () => {
    expect(sideLine('none', 'right', edge, LINE)).toEqual({ style: 'none' })
    expect(sideLine('all', 'right', edge, LINE)).toBe(LINE)
  })
})

describe('setSelectionBorders', () => {
  it('a bare caret formats the whole table (Word "Apply to: Table")', () => {
    const editor = createEditor()
    caretIn(editor, 1, 1)
    expect(run(editor, 'bottom')).toBe(true)
    expect(bordersAt(editor, 2, 0)?.bottom).toEqual(LINE)
    expect(bordersAt(editor, 2, 2)?.bottom).toEqual(LINE)
    expect(bordersAt(editor, 1, 1)).toBeNull()
    expect(bordersAt(editor, 0, 0)).toBeNull()
  })

  it('a cell selection only formats the selected block', () => {
    const editor = createEditor()
    selectCells(editor, [0, 0], [1, 1])
    expect(run(editor, 'outer')).toBe(true)
    expect(bordersAt(editor, 0, 0)).toEqual({ top: LINE, left: LINE })
    expect(bordersAt(editor, 1, 1)).toEqual({ bottom: LINE, right: LINE })
    expect(bordersAt(editor, 2, 2)).toBeNull()
  })

  it('inside horizontal paints the shared row edges and nothing on the outline', () => {
    const editor = createEditor()
    selectCells(editor, [0, 0], [2, 0])
    expect(run(editor, 'insideH')).toBe(true)
    expect(bordersAt(editor, 0, 0)).toEqual({ bottom: LINE })
    expect(bordersAt(editor, 1, 0)).toEqual({ top: LINE, bottom: LINE })
    expect(bordersAt(editor, 2, 0)).toEqual({ top: LINE })
    expect(bordersAt(editor, 0, 1)).toBeNull()
  })

  it('inside modes on a single cell change nothing', () => {
    const editor = createEditor()
    selectCells(editor, [1, 1], [1, 1])
    expect(run(editor, 'insideV')).toBe(false)
    expect(bordersAt(editor, 1, 1)).toBeNull()
  })

  it('No Border writes explicit none sides and drops table-level inside lines', () => {
    const inside = { style: 'single', szEighths: 4 }
    const editor = createEditor({ borders: { top: inside, insideH: inside, insideV: inside } })
    caretIn(editor, 0, 0)
    expect(run(editor, 'none')).toBe(true)
    expect(bordersAt(editor, 1, 1)).toEqual({
      top: { style: 'none' },
      bottom: { style: 'none' },
      left: { style: 'none' },
      right: { style: 'none' },
    })
    expect(editor.state.doc.child(1).attrs.borders).toEqual({ top: inside })
  })

  it('keeps unrelated sides a cell already had', () => {
    const editor = createEditor()
    selectCells(editor, [0, 0], [0, 0])
    run(editor, 'top')
    run(editor, 'left')
    expect(bordersAt(editor, 0, 0)).toEqual({ top: LINE, left: LINE })
  })

  it('refuses outside a table', () => {
    const editor = createEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)))
    expect(run(editor, 'all')).toBe(false)
  })
})
