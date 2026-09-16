import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { handleDocsControl } from '../src/renderer/control'

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const para = (text: string) => ({ type: 'docParagraph', content: [{ type: 'text', text }] })

function createEditor(texts: string[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: texts.map(para) },
  })
  editors.add(editor)
  return editor
}

describe('docs control hook', () => {
  it('waits for a document, then selects the requested block', () => {
    const editor = createEditor(['First paragraph.', 'Second paragraph.', 'Third.'])
    expect(handleDocsControl({ cmd: 'selection' }, editor, false)).toEqual({ status: 'not_ready' })
    expect(
      handleDocsControl({ cmd: 'goto', target: { kind: 'block', block: 1 } }, editor, true),
    ).toEqual({
      status: 'ok',
      result: { block: 1, type: 'docParagraph' },
    })
    const { from, to } = editor.state.selection
    expect(editor.state.doc.textBetween(from, to)).toBe('Second paragraph.')
    expect(
      handleDocsControl({ cmd: 'goto', target: { kind: 'block', block: 7 } }, editor, true),
    ).toMatchObject({
      status: 'error',
      error: { reason: 'out_of_range', detail: { valid_range: '0-2' } },
    })
  })

  it('reports the selected block range and text', () => {
    const editor = createEditor(['Alpha beta.', 'Gamma delta.', 'Epsilon.'])
    const block0 = editor.state.doc.child(0).nodeSize
    editor.commands.setTextSelection({ from: 3, to: block0 + 6 })
    expect(handleDocsControl({ cmd: 'selection' }, editor, true)).toEqual({
      status: 'ok',
      result: { blocks: [0, 1], text: 'pha beta.\nGamma', collapsed: false },
    })
    editor.commands.setTextSelection(block0 + 2)
    expect(handleDocsControl({ cmd: 'selection' }, editor, true)).toEqual({
      status: 'ok',
      result: { blocks: [1, 1], text: '', collapsed: true },
    })
    // a caret at the very end of the document still belongs to the last block
    editor.commands.setTextSelection(editor.state.doc.content.size)
    expect(handleDocsControl({ cmd: 'selection' }, editor, true)).toMatchObject({
      status: 'ok',
      result: { blocks: [2, 2], collapsed: true },
    })
    // a whole-paragraph selection ends on the next block's boundary and must not count it
    editor.commands.setTextSelection({ from: 1, to: block0 })
    expect(handleDocsControl({ cmd: 'selection' }, editor, true)).toEqual({
      status: 'ok',
      result: { blocks: [0, 0], text: 'Alpha beta.', collapsed: false },
    })
  })
})
