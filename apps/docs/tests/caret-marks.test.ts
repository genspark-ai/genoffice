/**
 * Arrowing away from a fresh empty paragraph and back must not drop the
 * pending format (Word keeps it on the paragraph mark): Calibri text +
 * Enter + Up + Down + type used to produce the document default (r114).
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const CALIBRI = { font: 'Calibri', sizeHalfPoints: 24 }

const editorWithStyledText = () =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            {
              type: 'text',
              text: 'hello',
              marks: [{ type: 'docTextStyle', attrs: CALIBRI }],
            },
          ],
        },
      ],
    } as never,
  })

const textStyleOf = (
  marks: readonly { type: { name: string }; attrs: Record<string, unknown> }[],
) => marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs

describe('caret marks memory (r114)', () => {
  it('keeps the pending format across arrow navigation out of and back into an empty paragraph', () => {
    const editor = editorWithStyledText()
    // caret at end of "hello", split like Enter does (marks carry as storedMarks)
    editor.commands.setTextSelection(6)
    editor.commands.splitBlock({ keepMarks: true })
    expect(editor.state.doc.childCount).toBe(2)
    // the plugin stamped the pending marks onto the empty paragraph
    const stamped = editor.state.doc.child(1).attrs.caretMarks as string | null
    expect(stamped).toContain('Calibri')

    // arrow up (selection into the first paragraph) clears stored marks
    editor.commands.setTextSelection(3)
    // arrow back down into the empty paragraph
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    const restored = editor.state.storedMarks
    expect(restored).not.toBeNull()
    expect(textStyleOf(restored!)).toMatchObject(CALIBRI)

    // typing consumes the restored marks
    editor.view.dispatch(editor.state.tr.insertText('x'))
    const typed = editor.state.doc.child(1).firstChild!
    expect(typed.text).toBe('x')
    expect(textStyleOf(typed.marks)).toMatchObject(CALIBRI)
    editor.destroy()
  })

  it('an explicitly cleared pending format is respected (empty storedMarks array wins)', () => {
    const editor = editorWithStyledText()
    editor.commands.setTextSelection(6)
    editor.commands.splitBlock({ keepMarks: true })
    // user clears formatting on the empty paragraph: storedMarks becomes []
    editor.view.dispatch(editor.state.tr.setStoredMarks([]))
    // stamp follows the clear
    expect(editor.state.doc.child(1).attrs.caretMarks).toBeNull()
    editor.destroy()
  })

  it('does not invent marks for paragraphs that never had a pending format', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph' }, { type: 'docParagraph' }],
      } as never,
    })
    editor.commands.setTextSelection(1)
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    expect(editor.state.storedMarks).toBeNull()
    editor.destroy()
  })
})
