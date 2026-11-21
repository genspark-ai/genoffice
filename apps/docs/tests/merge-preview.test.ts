/** Preview Results shows a recipient's values in the merge fields without touching the document. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { toggleAllFieldCodes } from '../src/renderer/editor/field-codes'
import { mergePreviewRecord, setMergePreview } from '../src/renderer/editor/merge-preview'

const field = (instr: string, text: string) => ({
  type: 'text',
  text,
  marks: [{ type: 'instrField', attrs: { instr } }],
})

describe('merge preview', () => {
  let editor: Editor
  const shown = () =>
    Array.from(editor.view.dom.querySelectorAll('.doc-merge-preview')).map((el) => el.textContent)
  const hidden = () =>
    Array.from(editor.view.dom.querySelectorAll('.doc-field-result-hidden')).map(
      (el) => el.textContent,
    )
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'Dear ' },
              field(' MERGEFIELD First ', '«First»'),
              { type: 'text', text: ', page ' },
              field(' PAGE ', '1'),
            ],
          },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  it('shows the record in the merge fields, leaves other fields alone, and goes back', () => {
    const before = editor.state.doc
    setMergePreview(editor, { First: 'Ada' })
    expect(shown()).toEqual(['Ada'])
    expect(hidden()).toEqual(['«First»'])
    expect(mergePreviewRecord(editor.state)).toEqual({ First: 'Ada' })
    // a display only: the document is the one it was
    expect(editor.state.doc.eq(before)).toBe(true)
    expect(editor.can().undo()).toBe(false)

    setMergePreview(editor, { First: 'Grace' })
    expect(shown()).toEqual(['Grace'])
    setMergePreview(editor, null)
    expect(shown()).toEqual([])
    expect(hidden()).toEqual([])
  })

  it('follows the document: a field typed in while previewing shows its value too', () => {
    setMergePreview(editor, { First: 'Ada', Last: 'Lovelace' })
    editor.commands.insertContentAt(6, [field(' MERGEFIELD Last ', '«Last»')] as never)
    expect(shown().sort()).toEqual(['Ada', 'Lovelace'])
  })

  it("draws the value in the field's own formatting", () => {
    editor.commands.setTextSelection({ from: 6, to: 13 })
    editor.commands.setMark('bold')
    setMergePreview(editor, { First: 'Ada' })
    const value = editor.view.dom.querySelector('.doc-merge-preview')!
    expect(value.closest('strong, b')).not.toBeNull()
  })

  it('steps aside for a field that is showing its code, and comes back when the code is hidden', () => {
    setMergePreview(editor, { First: 'Ada' })
    toggleAllFieldCodes(editor.view)
    expect(shown()).toEqual([])
    expect(editor.view.dom.querySelectorAll('.doc-field-code')).toHaveLength(2)
    toggleAllFieldCodes(editor.view)
    expect(shown()).toEqual(['Ada'])
  })
})
