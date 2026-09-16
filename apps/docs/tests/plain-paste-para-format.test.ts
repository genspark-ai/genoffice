/**
 * Multi-line plain-text paste keeps the insertion paragraph's PARAGRAPH
 * formatting on every created paragraph (Word Keep Text Only formats like
 * typing Enter there). Only the first line merges into the destination
 * paragraph, so lines 2+ landed as DEFAULT paragraphs and dropped the
 * surrounding indent/line-spacing/alignment (user report: multi-
 * paragraph Ctrl+Shift+V applied the document paragraph format to the
 * first paragraph only).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pasteTextSlice } from '../src/renderer/editor/paste-text'

class FakeClipboardEvent extends Event {
  clipboardData = null
}
;(globalThis as Record<string, unknown>).ClipboardEvent = FakeClipboardEvent

const PARA_FORMAT = {
  align: 'both',
  lineSpacing: 2,
  indentLeft: 240,
  indentFirstLine: 0,
  styleId: 'Normal',
}

function makeEditor(paraAttrs: Record<string, unknown> = PARA_FORMAT) {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    editorProps: {
      clipboardTextParser: (text, $context, _plain, view) => pasteTextSlice(text, $context, view),
    },
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: paraAttrs,
          content: [{ type: 'text', text: 'Existing paragraph text' }],
        },
      ],
    },
  })
}

const HOST_LEN = 'Existing paragraph text'.length

function caretToEnd(editor: Editor) {
  const { state, dispatch } = editor.view
  dispatch(state.tr.setSelection(TextSelection.create(state.doc, 1 + HOST_LEN)))
}

const paraAttrs = (editor: Editor, index: number) => editor.state.doc.child(index).attrs

describe('multi-line plain paste keeps the insertion paragraph format', () => {
  it('every created paragraph clones the host paragraph formatting', () => {
    const editor = makeEditor()
    caretToEnd(editor)
    editor.view.pasteText('first line\nsecond line\nthird line')
    expect(editor.state.doc.childCount).toBe(3)
    for (let i = 0; i < 3; i++) expect(paraAttrs(editor, i)).toMatchObject(PARA_FORMAT)
    editor.destroy()
  })

  it('mid-paragraph paste formats the split tail like the host too', () => {
    const editor = makeEditor()
    const { state, dispatch } = editor.view
    // caret between "Existing" and " paragraph text"
    dispatch(state.tr.setSelection(TextSelection.create(state.doc, 1 + 'Existing'.length)))
    editor.view.pasteText('alpha\nbeta')
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).textContent).toBe('beta paragraph text')
    expect(paraAttrs(editor, 0)).toMatchObject(PARA_FORMAT)
    expect(paraAttrs(editor, 1)).toMatchObject(PARA_FORMAT)
    editor.destroy()
  })

  it('identity attrs are never cloned onto created paragraphs', () => {
    const editor = makeEditor({ ...PARA_FORMAT, docxIndex: 7 })
    caretToEnd(editor)
    editor.view.pasteText('one\ntwo')
    expect(paraAttrs(editor, 0).docxIndex).toBe(7)
    expect(paraAttrs(editor, 1).docxIndex).toBeNull()
    expect(paraAttrs(editor, 1)).toMatchObject(PARA_FORMAT)
    editor.destroy()
  })

  it('a default host paragraph still produces default paragraphs', () => {
    const editor = makeEditor({})
    caretToEnd(editor)
    editor.view.pasteText('one\ntwo')
    expect(paraAttrs(editor, 1).align).toBeNull()
    expect(paraAttrs(editor, 1).lineSpacing).toBeNull()
    editor.destroy()
  })
})
