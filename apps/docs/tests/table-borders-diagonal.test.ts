import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  setSelectionBorders,
  sideLine,
  type BorderLine,
  type TableBorderMode,
} from '../src/renderer/editor/table-borders'
import { cellDiagonalCss } from '../src/renderer/editor/border-metrics'

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

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('Intro'), table(), para('Outro')] },
  })
  editors.add(editor)
  return editor
}

const LINE: BorderLine = { style: 'single', szEighths: 12, color: 'FF0000' }

function cellPositions(editor: Editor): number[][] {
  const out: number[][] = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTableRow') out.push([])
    if (node.type.name === 'docTableCell') out[out.length - 1].push(pos)
    return true
  })
  return out
}

function bordersAt(editor: Editor, row: number, col: number): Record<string, BorderLine> {
  const pos = cellPositions(editor)[row][col]
  return (editor.state.doc.nodeAt(pos)!.attrs.borders ?? {}) as Record<string, BorderLine>
}

function caretIn(editor: Editor, row: number, col: number) {
  const pos = cellPositions(editor)[row][col]
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 2)),
  )
}

function run(editor: Editor, mode: TableBorderMode): boolean {
  return setSelectionBorders(mode, LINE)(editor.state, editor.view.dispatch)
}

const EDGE = { top: true, bottom: true, left: true, right: true }

describe('sideLine: the diagonals', () => {
  it('applies only the diagonal the gallery entry names', () => {
    expect(sideLine('tl2br', 'tl2br', EDGE, LINE)).toEqual(LINE)
    expect(sideLine('tl2br', 'tr2bl', EDGE, LINE)).toBeUndefined()
    expect(sideLine('tr2bl', 'tr2bl', EDGE, LINE)).toEqual(LINE)
    expect(sideLine('tr2bl', 'tl2br', EDGE, LINE)).toBeUndefined()
  })

  it('a diagonal ignores the inner/outer test, since it is never a shared edge', () => {
    const inner: Record<'top' | 'bottom' | 'left' | 'right', boolean> = {
      top: false,
      bottom: false,
      left: false,
      right: false,
    }
    expect(sideLine('tl2br', 'tl2br', inner, LINE)).toEqual(LINE)
  })

  it('No Border clears a diagonal, matching the four edges', () => {
    expect(sideLine('none', 'tl2br', EDGE, LINE)).toEqual({ style: 'none' })
    expect(sideLine('none', 'tr2bl', EDGE, LINE)).toEqual({ style: 'none' })
  })

  it('the compound modes stay edges-only, as in Word', () => {
    for (const mode of ['all', 'outer', 'inner'] as const) {
      expect(sideLine(mode, 'tl2br', EDGE, LINE)).toBeUndefined()
      expect(sideLine(mode, 'tr2bl', EDGE, LINE)).toBeUndefined()
    }
  })
})

describe('setSelectionBorders: the diagonals', () => {
  it('stamps the diagonal on every targeted cell without touching the edges', () => {
    const editor = createEditor()
    caretIn(editor, 0, 0)
    expect(run(editor, 'tl2br')).toBe(true)
    for (let r = 0; r < 3; r++) {
      for (let c = 0; c < 3; c++) {
        expect(bordersAt(editor, r, c)).toEqual({ tl2br: LINE })
      }
    }
  })

  it('a cell selection limits the diagonal to the selected cells', () => {
    const editor = createEditor()
    const cells = cellPositions(editor)
    editor.view.dispatch(
      editor.state.tr.setSelection(
        CellSelection.create(editor.state.doc, cells[0][0], cells[0][1]),
      ),
    )
    expect(run(editor, 'tr2bl')).toBe(true)
    expect(bordersAt(editor, 0, 0)).toEqual({ tr2bl: LINE })
    expect(bordersAt(editor, 0, 1)).toEqual({ tr2bl: LINE })
    expect(bordersAt(editor, 1, 1)).toEqual({})
  })

  it('keeps both diagonals when they are set in turn', () => {
    const editor = createEditor()
    caretIn(editor, 1, 1)
    run(editor, 'tl2br')
    run(editor, 'tr2bl')
    expect(bordersAt(editor, 1, 1)).toEqual({ tl2br: LINE, tr2bl: LINE })
  })

  it('No Border removes an existing diagonal, along with the four edges', () => {
    const editor = createEditor()
    caretIn(editor, 0, 0)
    run(editor, 'tl2br')
    expect(run(editor, 'none')).toBe(true)
    expect(bordersAt(editor, 0, 0)).toEqual({
      top: { style: 'none' },
      bottom: { style: 'none' },
      left: { style: 'none' },
      right: { style: 'none' },
      tl2br: { style: 'none' },
      tr2bl: { style: 'none' },
    })
  })

  it('All Borders leaves an existing diagonal alone, as in Word', () => {
    const editor = createEditor()
    caretIn(editor, 0, 0)
    run(editor, 'tl2br')
    run(editor, 'all')
    expect(bordersAt(editor, 0, 0).tl2br).toEqual(LINE)
  })
})

describe('cellDiagonalCss', () => {
  it('aims each diagonal at the matching CSS magic corner', () => {
    // a hard-stop band on `to bottom right` runs along the top-left to
    // bottom-right diagonal at any aspect ratio; `to top right` is the other one
    expect(cellDiagonalCss({ style: 'single', szEighths: 8 }, 'tl2br')).toContain('to bottom right')
    expect(cellDiagonalCss({ style: 'single', szEighths: 8 }, 'tr2bl')).toContain('to top right')
  })

  it('bands at half the drawn width on each side of the centre', () => {
    // sz 8 (1pt) draws at borderDrawnPx's rounded thickness, so the stops
    // straddle the centre by half of it
    const css = cellDiagonalCss({ style: 'single', szEighths: 8 }, 'tl2br')!
    expect(css).toContain('calc(50% - 0.5px)')
    expect(css).toContain('calc(50% + 0.5px)')
    // a hairline never collapses below the 1px the four edges also clamp to
    expect(cellDiagonalCss({ style: 'single', szEighths: 2 }, 'tl2br')).toContain('0.5px')
  })

  it('honours the authored color and falls back to black for auto', () => {
    expect(cellDiagonalCss({ style: 'single', color: 'FF0000' }, 'tl2br')).toContain('#FF0000')
    expect(cellDiagonalCss({ style: 'single', color: 'auto' }, 'tl2br')).toContain('#000')
  })

  it('draws nothing for a missing, none or nil line', () => {
    expect(cellDiagonalCss(undefined, 'tl2br')).toBeNull()
    expect(cellDiagonalCss({ style: 'none' }, 'tl2br')).toBeNull()
    expect(cellDiagonalCss({ style: 'nil' }, 'tr2bl')).toBeNull()
  })
})
