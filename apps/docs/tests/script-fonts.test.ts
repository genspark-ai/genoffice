import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: text ? [{ type: 'text', text }] : [],
})

function open(paragraphs: string[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: paragraphs.map(para) },
  })
  editors.push(editor)
  return editor
}

const decorated = (editor: Editor) =>
  [...editor.view.dom.querySelectorAll('.doc-east-asian-font')].map((el) => el.textContent)

describe('ScriptFonts decorations', () => {
  it('wraps only the East Asian runs', () => {
    const editor = open(['Report \u62a5\u544a 2026', 'plain', '\u4e2d\u6587'])
    expect(decorated(editor)).toEqual(['\u62a5\u544a', '\u4e2d\u6587'])
  })

  it('tracks an edit in one block without losing the others', () => {
    const editor = open(['a \u7532 b', 'x', 'c \u4e59 d'])
    // typing into the middle block: earlier ranges map through, later ones too
    editor.commands.insertContentAt(editor.state.doc.child(0).nodeSize + 2, '\u4e19')
    expect(decorated(editor)).toEqual(['\u7532', '\u4e19', '\u4e59'])
    // removing the character drops its decoration again
    const pos = editor.state.doc.child(0).nodeSize + 2
    editor.commands.deleteRange({ from: pos, to: pos + 1 })
    expect(decorated(editor)).toEqual(['\u7532', '\u4e59'])
  })

  it('matches a full rebuild after many-block edits', () => {
    const editor = open(Array.from({ length: 6 }, (_, i) => `p${i} \u6587`))
    editor.commands.selectAll()
    editor.commands.insertContent('\u65b0 new')
    expect(decorated(editor)).toEqual(['\u65b0'])
  })
})
