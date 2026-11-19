import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildDocumentContext, countWords } from '../src/renderer/ai/protocol'
import { documentTextForWordCount } from '../src/renderer/word-count'

/**
 * The AI snapshot reports its own "Full-text stats" line, and office_doc_info
 * reports the status-bar figure. The two disagreed whenever blocks met without
 * whitespace (their last and first words merged) or when a formula/ruby atom
 * carried text the snapshot walked past. Both now derive the count from the
 * same rule; this pins the parity.
 */

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})
const math = (t: string): JsonNode => ({
  type: 'docInlineMath',
  attrs: { omml: '', mathml: '', latex: t, text: t },
})

const liveEditors: Editor[] = []

afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}

describe('AI snapshot word count vs office_doc_info', () => {
  it('separates adjacent blocks and counts inline formula text', () => {
    // "alpha" + "beta" concatenated would be one word; the formula adds two
    // more. The right answer is 4.
    const editor = createEditor([para(text('alpha')), para(text('beta')), para(math('E=mc2'))])
    const infoWords = countWords(documentTextForWordCount(editor.state.doc))
    expect(infoWords).toBe(4)
    expect(buildDocumentContext(editor)).toContain(`Full-text stats: words ${infoWords}`)
  })

  it('still agrees on ordinary prose with punctuation and CJK', () => {
    const editor = createEditor([
      para(text('Hello, world.')),
      para(text('\u7B2C\u4E8C\u6BB5\u6587\u5B57')),
      para(text('Third block ends with Latin')),
    ])
    const infoWords = countWords(documentTextForWordCount(editor.state.doc))
    expect(buildDocumentContext(editor)).toContain(`Full-text stats: words ${infoWords}`)
  })

  it('agrees across tables and lists, down to the character counts', () => {
    const cell = (t: string): JsonNode => ({
      type: 'docTableCell',
      content: [para(text(t))],
    })
    const table: JsonNode = {
      type: 'docTable',
      content: [{ type: 'docTableRow', content: [cell('beta'), cell('gamma')] }],
    }
    const li = (t: string): JsonNode => ({
      type: 'docListItem',
      attrs: { docxIndex: null, kind: 'bullet', numId: null, ilvl: 0 },
      content: [text(t)],
    })
    const editor = createEditor([
      para(text('alpha')),
      table,
      li('delta'),
      li('epsilon'),
      para(math('E=mc2')),
    ])
    const docText = documentTextForWordCount(editor.state.doc)
    expect(countWords(docText)).toBe(7)
    const stats = `Full-text stats: words 7, characters (no spaces) ${
      docText.replace(/\s/g, '').length
    }, characters (with spaces) ${docText.length}`
    expect(buildDocumentContext(editor)).toContain(stats)
  })
})
