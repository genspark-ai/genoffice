import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyPainterFormat, type PainterFormat } from '../src/renderer/components/Ribbon'

const para = (text: string) => ({
  type: 'docParagraph',
  attrs: {},
  content: [{ type: 'text', text }],
})

const HEADING: PainterFormat = {
  marks: [],
  block: { type: 'docHeading', attrs: { level: 1, styleId: 'Heading1' } },
}

const types = (editor: Editor) => editor.state.doc.content.content.map((n) => n.type.name)

describe('format painter undo grouping', () => {
  it('undoes quick strokes on neighbouring paragraphs one at a time', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [para('alpha one'), para('beta two'), para('gamma')],
      } as never,
    })
    applyPainterFormat(editor, HEADING, 1, 6, 1)
    applyPainterFormat(editor, HEADING, 12, 16, 12)
    expect(types(editor)).toEqual(['docHeading', 'docHeading', 'docParagraph'])
    editor.commands.undo()
    expect(types(editor)).toEqual(['docHeading', 'docParagraph', 'docParagraph'])
    editor.commands.undo()
    expect(types(editor)).toEqual(['docParagraph', 'docParagraph', 'docParagraph'])
    editor.destroy()
  })
})
