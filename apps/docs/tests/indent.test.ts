import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { stepParagraphIndent } from '../src/renderer/editor/indent'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (t: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content: [text(t)],
})

function createEditor(content: JsonNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

function select(editor: Editor, from: number, to = from) {
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
  )
}

const indentOf = (editor: Editor, index: number) => editor.state.doc.child(index).attrs.indentLeft

describe('stepParagraphIndent', () => {
  it('steps a plain paragraph by half-inch stops and back to null', () => {
    const editor = createEditor([para('plain paragraph')])
    select(editor, 2)
    expect(stepParagraphIndent(editor, 1)).toBe(true)
    expect(indentOf(editor, 0)).toBe(720)
    stepParagraphIndent(editor, 1)
    expect(indentOf(editor, 0)).toBe(1440)
    stepParagraphIndent(editor, -1)
    expect(indentOf(editor, 0)).toBe(720)
    stepParagraphIndent(editor, -1)
    expect(indentOf(editor, 0)).toBeNull()
    editor.destroy()
  })

  it('snaps odd indents to the surrounding stops', () => {
    const editor = createEditor([para('para', { indentLeft: 500 })])
    select(editor, 2)
    stepParagraphIndent(editor, 1)
    expect(indentOf(editor, 0)).toBe(720)
    const editor2 = createEditor([para('para', { indentLeft: 500 })])
    select(editor2, 2)
    stepParagraphIndent(editor2, -1)
    expect(indentOf(editor2, 0)).toBeNull()
    editor.destroy()
    editor2.destroy()
  })

  it('does not go below zero', () => {
    const editor = createEditor([para('para')])
    select(editor, 2)
    expect(stepParagraphIndent(editor, -1)).toBe(false)
    expect(indentOf(editor, 0)).toBeNull()
    editor.destroy()
  })

  it('indents every paragraph and heading in the selection', () => {
    const editor = createEditor([
      para('first paragraph'),
      {
        type: 'docHeading',
        attrs: { docxIndex: null, level: 2 },
        content: [text('heading')],
      },
      para('third paragraph'),
    ])
    select(editor, 2, editor.state.doc.content.size - 2)
    stepParagraphIndent(editor, 1)
    expect(indentOf(editor, 0)).toBe(720)
    expect(indentOf(editor, 1)).toBe(720)
    expect(indentOf(editor, 2)).toBe(720)
    editor.destroy()
  })

  it('changes list level instead of indentLeft for list items', () => {
    const editor = createEditor([
      {
        type: 'docListItem',
        attrs: { docxIndex: null, kind: 'bullet', ilvl: 0 },
        content: [text('item')],
      },
    ])
    select(editor, 2)
    stepParagraphIndent(editor, 1)
    expect(editor.state.doc.child(0).attrs.ilvl).toBe(1)
    expect(indentOf(editor, 0)).toBeNull()
    stepParagraphIndent(editor, -1)
    expect(editor.state.doc.child(0).attrs.ilvl).toBe(0)
    editor.destroy()
  })
})
