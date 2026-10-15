import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { selectedWordCount } from '../src/renderer/word-count'

// Word shows "selected/total" while a selection exists and the plain total
// once the caret collapses again.

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [text(t)],
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

describe('selectedWordCount', () => {
  it('is 0 for a collapsed caret, so the status bar shows the plain total', () => {
    const editor = createEditor([para('\u601d\u7ef4\u5bfc\u56fe\u5df2\u7ecf\u63d0\u4ea4')])
    editor.commands.setTextSelection(3)
    expect(selectedWordCount(editor.state.doc, editor.state.selection)).toBe(0)
  })

  it('counts the selection with the same Word rule as the total', () => {
    const editor = createEditor([para('\u601d\u7ef4\u5bfc\u56fe\u5df2\u7ecf\u63d0\u4ea4')])
    // select the first 4 asian chars (counted one by one)
    editor.commands.setTextSelection({ from: 1, to: 5 })
    expect(selectedWordCount(editor.state.doc, editor.state.selection)).toBe(4)
  })

  it('counts a mixed CJK/latin selection the way the dialog would', () => {
    const editor = createEditor([para('GenOffice \u601d\u7ef4\u5bfc\u56fe')])
    // the whole paragraph: "GenOffice" as 1 word + 4 asian chars
    editor.commands.setTextSelection({ from: 1, to: editor.state.doc.content.size - 1 })
    expect(selectedWordCount(editor.state.doc, editor.state.selection)).toBe(5)
  })

  it('sums a selection that spans several paragraphs', () => {
    const editor = createEditor([
      para('\u601d\u7ef4\u5bfc\u56fe'),
      para('\u5df2\u7ecf\u5408\u5e76'),
    ])
    // the whole document: both paragraphs, 4 + 4 asian chars
    editor.commands.setTextSelection({
      from: 1,
      to: editor.state.doc.content.size - 1,
    })
    expect(selectedWordCount(editor.state.doc, editor.state.selection)).toBe(8)
  })

  it('a selection of only whitespace has no word figure', () => {
    // the blank-paragraph case: textBetween returns "   ", which is not empty
    // but carries no words. Counting it must still read 0, or the status bar
    // would flash "3 / 900 words" while the user drags across a gap.
    const editor = createEditor([para('   ')])
    editor.commands.setTextSelection({ from: 1, to: 4 })
    expect(selectedWordCount(editor.state.doc, editor.state.selection)).toBe(0)
  })

  it('a NodeSelection (an image) has no text range, so the figure is 0', () => {
    const editor = createEditor([
      { type: 'docImage', attrs: { docxIndex: null, src: 'data:image/png;base64,AA' } },
    ])
    // an image is selected as a node, not as text: from/to sit around it and
    // textBetween() yields ''. A helper that assumed a text range would either
    // crash here or report the document's own words as "selected".
    const sel = NodeSelection.create(editor.state.doc, 0)
    expect(sel.empty).toBe(false)
    expect(selectedWordCount(editor.state.doc, sel)).toBe(0)
  })
})
