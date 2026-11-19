import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

const heading = { type: 'docHeading', attrs: { level: 1, styleId: 'Heading1' } }

function editorWith(text: string): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ ...heading, content: [{ type: 'text', text }] }],
    } as never,
  })
}

const pressEnter = (editor: Editor) =>
  editor.view.dom.dispatchEvent(
    new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true, cancelable: true }),
  )

const shape = (editor: Editor) =>
  editor.state.doc.content.content.map((n) => `${n.type.name}:${n.attrs.styleId}:${n.textContent}`)

describe('Enter at the end of a heading', () => {
  it('starts a body paragraph without the heading style (Word: next = Normal)', () => {
    const editor = editorWith('Intro')
    editor.commands.setTextSelection(6)
    pressEnter(editor)
    expect(shape(editor)).toEqual(['docHeading:Heading1:Intro', 'docParagraph:null:'])
    editor.destroy()
  })

  it('leaves a page-break-before heading in place and does not clone the break', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docHeading',
            attrs: { level: 1, styleId: 'Heading1', pageBreakBefore: true },
            content: [{ type: 'text', text: 'Intro' }],
          },
        ],
      } as never,
    })
    editor.commands.setTextSelection(6)
    pressEnter(editor)
    expect(shape(editor)).toEqual(['docHeading:Heading1:Intro', 'docParagraph:null:'])
    expect(editor.state.doc.child(0).attrs.pageBreakBefore).toBe(true)
    expect(editor.state.doc.child(1).attrs.pageBreakBefore).toBe(false)
    editor.destroy()
  })

  it('keeps the heading style on both halves of a mid-heading split', () => {
    const editor = editorWith('Intro')
    editor.commands.setTextSelection(3)
    pressEnter(editor)
    expect(shape(editor)).toEqual(['docHeading:Heading1:In', 'docHeading:Heading1:tro'])
    editor.destroy()
  })
})
