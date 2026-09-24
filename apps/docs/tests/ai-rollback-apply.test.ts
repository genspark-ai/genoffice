import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyDocument } from '../src/renderer/ai/rollback-doc'
import type { PmNode } from '../src/renderer/editor/convert'

// A live view leaves a DOMObserver flush timer behind; destroy every editor or
// vitest reports an unhandled ReferenceError after the environment tears down.
const liveEditors: Editor[] = []
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

function para(text: string): PmNode {
  return { type: 'docParagraph', attrs: { docxIndex: null }, content: [{ type: 'text', text }] }
}

function doc(...texts: string[]): PmNode {
  return { type: 'doc', content: texts.map(para) }
}

function createEditor(initial: PmNode): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: initial as never,
  })
  liveEditors.push(editor)
  return editor
}

const textOf = (editor: Editor) =>
  editor.state.doc.textBetween(0, editor.state.doc.content.size, '\n')

describe('applyDocument', () => {
  it('replaces the whole document with the version snapshot', () => {
    const editor = createEditor(doc('before the turn'))
    applyDocument(editor, doc('as it was before', 'two paragraphs'))
    expect(textOf(editor)).toBe('as it was before\ntwo paragraphs')
  })

  it('keeps the swap out of the undo history', () => {
    // #543: the panel owns roll back and its undo. If this swap landed in the
    // history, `⌘Z` could put the document back while the version list still
    // said the roll back stood — the exact drift the issue is about.
    const editor = createEditor(doc('original'))

    // the harness's own undo works — otherwise the assertions below prove nothing
    editor.commands.insertContent(' typed')
    expect(editor.commands.undo()).toBe(true)
    expect(textOf(editor)).toBe('original')

    applyDocument(editor, doc('rolled back'))
    expect(textOf(editor)).toBe('rolled back')
    // nothing left to undo: the swap itself was not recorded as an undo step
    expect(editor.commands.undo()).toBe(false)
    expect(textOf(editor)).toBe('rolled back')
  })

  it('is safe for an empty document', () => {
    const editor = createEditor(doc('something'))
    applyDocument(editor, { type: 'doc', content: [] })
    expect(textOf(editor)).toBe('')
  })
})
