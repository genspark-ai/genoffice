/**
 * Issue #126 wave 2: paragraph-style and line-spacing commands backing the
 * ⌥⌘0-3 / ⌘1·2·5 shortcuts. applyParagraphStyle was extracted from the ribbon
 * gallery closure; these pin its Word-like behaviors (node switch + shedding
 * the runs' direct font/size/color).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyParagraphStyle, setParaAttrs } from '../src/renderer/components/ribbon-tabs'

const styled = { type: 'docTextStyle', attrs: { sizeHalfPoints: 48, color: 'FF0000' } }

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'chapter title', marks: [styled] }],
        },
        { type: 'docParagraph', content: [{ type: 'text', text: 'body text' }] },
      ],
    },
  })
}

describe('applyParagraphStyle', () => {
  it('switches the block to a heading and sheds direct size/color', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h2')
    const block = editor.state.doc.child(0)
    expect(block.type.name).toBe('docHeading')
    expect(block.attrs.level).toBe(2)
    const marks = block.firstChild!.marks.filter((m) => m.type.name === 'docTextStyle')
    expect(marks.length).toBe(0)
    editor.destroy()
  })

  it('returns a heading to a normal paragraph', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3)))
    applyParagraphStyle(editor, 'h1')
    applyParagraphStyle(editor, 'p')
    expect(editor.state.doc.child(0).type.name).toBe('docParagraph')
    editor.destroy()
  })
})

describe('line spacing via setParaAttrs (⌘1/⌘2/⌘5 path)', () => {
  it('sets the multiple on every paragraph in the selection and clears exact rules', () => {
    const editor = makeEditor()
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3, 20)),
    )
    setParaAttrs(editor, { lineSpacing: 1.5, lineRule: null, lineRawTwips: null })
    expect(editor.state.doc.child(0).attrs.lineSpacing).toBe(1.5)
    expect(editor.state.doc.child(1).attrs.lineSpacing).toBe(1.5)
    expect(editor.state.doc.child(0).attrs.lineRule).toBeNull()
    editor.destroy()
  })
})
