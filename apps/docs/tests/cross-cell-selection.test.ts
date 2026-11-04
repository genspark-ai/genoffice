/**
 * Word-style continuous selection across table cells (genoffice#2329): the
 * boundary helpers, and the delete/type paths that clear the covered text per
 * cell while every cell they crossed survives.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { parseDocx } from '@genoffice/docx-engine'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  cellDeletionSpans,
  deleteCrossCellSelection,
  selectionCrossesCellBoundary,
} from '../src/renderer/editor/cross-cell-drag'

const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>Bravo</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>Charlie</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>Delta</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

async function openTable(): Promise<Editor> {
  const source = await buildDocx({ bodyXml: TABLE })
  const parsed = await parseDocx(source)
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

/** text positions by cell text, reading order: Alpha Bravo Charlie Delta */
function textStarts(editor: Editor): Map<string, number> {
  const starts = new Map<string, number>()
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.text && !starts.has(node.text)) starts.set(node.text, pos)
  })
  return starts
}

function cellTexts(editor: Editor): string[] {
  const texts: string[] = []
  editor.state.doc.descendants((node) => {
    const role = node.type.spec.tableRole
    if (role === 'cell' || role === 'header_cell') texts.push(node.textContent)
  })
  return texts
}

const keyEvent = (key: string): KeyboardEvent =>
  new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })

describe('selectionCrossesCellBoundary', () => {
  it('says which selections cross a cell boundary', async () => {
    const editor = await openTable()
    const s = textStarts(editor)
    const doc = editor.state.doc
    // inside one cell, even across that cell's paragraphs, is not cross-cell
    expect(selectionCrossesCellBoundary(doc, s.get('Alpha')!, s.get('Alpha')! + 3)).toBe(false)
    expect(selectionCrossesCellBoundary(doc, s.get('Bravo')!, s.get('Bravo')! + 2)).toBe(false)
    // one cell into another is
    expect(selectionCrossesCellBoundary(doc, s.get('Alpha')! + 2, s.get('Delta')! + 2)).toBe(true)
    // body text into a cell, and a cell into body text, cross too
    expect(selectionCrossesCellBoundary(doc, s.get('Alpha')! - 2, s.get('Bravo')! + 1)).toBe(true)
    expect(selectionCrossesCellBoundary(doc, s.get('Delta')! + 1, s.get('Delta')! - 2)).toBe(true)
    editor.destroy()
  })
})

describe('cellDeletionSpans', () => {
  it('covers the partial ends and the fully covered cells between', async () => {
    const editor = await openTable()
    const s = textStarts(editor)
    const spans = cellDeletionSpans(editor.state.doc, s.get('Alpha')! + 2, s.get('Delta')! + 2)
    expect(spans.length).toBe(4)
    // Alpha loses 'pha', Bravo and Charlie whole, Delta loses 'De'
    expect(editor.state.doc.textBetween(spans[0].from, spans[0].to)).toBe('pha')
    expect(editor.state.doc.textBetween(spans[1].from, spans[1].to)).toBe('Bravo')
    expect(editor.state.doc.textBetween(spans[2].from, spans[2].to)).toBe('Charlie')
    expect(editor.state.doc.textBetween(spans[3].from, spans[3].to)).toBe('De')
    editor.destroy()
  })
})

describe('deleting across cells', () => {
  it('clears the covered text per cell and keeps every cell', async () => {
    const editor = await openTable()
    const s = textStarts(editor)
    const from = s.get('Alpha')! + 2
    const to = s.get('Delta')! + 2
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
    )
    const handled = editor.view.someProp('handleKeyDown', (f) =>
      f(editor.view, keyEvent('Backspace')),
    )
    expect(handled).toBe(true)
    // the structure survived: still a 2x2 table, partial ends kept
    expect(cellTexts(editor)).toEqual(['Al', '', '', 'lta'])
    expect(editor.state.selection.from).toBe(from)
    editor.destroy()
  })

  it('keeps the default delete for selections inside one cell', async () => {
    const editor = await openTable()
    const s = textStarts(editor)
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, s.get('Alpha')!, s.get('Alpha')! + 2),
      ),
    )
    expect(deleteCrossCellSelection(editor.view)).toBe(false)
    editor.destroy()
  })

  it('does nothing for a collapsed caret', async () => {
    const editor = await openTable()
    const s = textStarts(editor)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, s.get('Bravo')!)),
    )
    expect(deleteCrossCellSelection(editor.view)).toBe(false)
    editor.destroy()
  })
})

describe('typing over a cross-cell selection', () => {
  it('lands the text where the selection began and keeps the structure', async () => {
    const editor = await openTable()
    const s = textStarts(editor)
    const from = s.get('Alpha')! + 2
    const to = s.get('Delta')! + 2
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
    )
    const handled = editor.view.someProp('handleTextInput', (f) =>
      f(editor.view, from, to, 'X', () => editor.state.tr),
    )
    expect(handled).toBe(true)
    expect(cellTexts(editor)).toEqual(['AlX', '', '', 'lta'])
    expect(editor.state.selection.from).toBe(from + 1)
    editor.destroy()
  })
})

describe('an IME starting over a cross-cell selection', () => {
  it('clears the covered text first so the composition begins on a collapsed caret', async () => {
    const editor = await openTable()
    const s = textStarts(editor)
    const from = s.get('Alpha')! + 2
    const to = s.get('Delta')! + 2
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
    )
    editor.view.dom.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    expect(cellTexts(editor)).toEqual(['Al', '', '', 'lta'])
    expect(editor.state.selection.from).toBe(from)
    editor.destroy()
  })
})
