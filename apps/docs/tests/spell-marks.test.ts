/** The red lines under misspelled words are the page's own: which words get one. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  refreshSpellMarks,
  setSpellMarksActive,
  spellMarkWords,
} from '../src/renderer/editor/spell-marks'
import {
  forgetIgnoredWords,
  ignoreSpellingOnce,
  setWordAside,
} from '../src/renderer/editor/spell-menu'

const WRONG = new Set(['teh', 'teeh', 'recieve', 'wrnog'])
const wrong = (word: string) => WRONG.has(word)

describe('spelling marks', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'teh cat and teh dog' }] },
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'to recieve ' },
              { type: 'text', text: 'wrnog', marks: [{ type: 'del', attrs: {} }] },
            ],
          },
        ],
      },
    })
  })
  afterEach(() => {
    forgetIgnoredWords()
    editor.destroy()
  })

  const marked = (from = 0, to = editor.state.doc.content.size) =>
    spellMarkWords(editor.state, from, to, wrong).map((w) => `${w.word}@${w.from}`)

  it('marks nothing until Spelling is on, then every misspelled word that is checked', () => {
    expect(marked()).toEqual([])
    setSpellMarksActive(editor, true)
    // text that came with the document, never typed next to; the deleted word is not checked
    expect(marked()).toEqual(['teh@1', 'teh@13', 'recieve@25'])
    setSpellMarksActive(editor, false)
    expect(marked()).toEqual([])
  })

  it('reads only the paragraphs between the two positions', () => {
    setSpellMarksActive(editor, true)
    expect(marked(0, 10)).toEqual(['teh@1'])
    expect(marked(22, 40)).toEqual(['recieve@25'])
  })

  it('leaves the word being typed alone until the caret is out of it', () => {
    setSpellMarksActive(editor, true)
    editor.commands.setTextSelection(20)
    for (const key of ' wrnog') editor.commands.insertContent(key)
    expect(marked()).toEqual(['teh@1', 'teh@13', 'recieve@31'])
    editor.commands.insertContent(' ')
    expect(marked()).toEqual(['teh@1', 'teh@13', 'wrnog@21', 'recieve@32'])
    // a caret put into a word that is already marked does not take its line away
    editor.commands.setTextSelection(2)
    expect(marked()).toContain('teh@1')
    // nor does a change that is not typing: formatting the word, or an edit elsewhere
    editor.commands.setTextSelection({ from: 1, to: 4 })
    editor.commands.setMark('bold')
    editor.commands.setTextSelection(2)
    expect(marked()).toContain('teh@1')
    editor.view.dispatch(editor.state.tr.insertText('x ', 14))
    expect(marked()).toContain('teh@1')
    // typing in it does, until the caret is out again
    editor.commands.insertContent('e')
    expect(marked()).not.toContain('teeh@1')
    editor.commands.setTextSelection(8)
    expect(marked()).toContain('teeh@1')
  })

  it('marks the word the caret lands in after the document is replaced or several words are pasted', () => {
    setSpellMarksActive(editor, true)
    editor.commands.setContent({
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'teh start' }] }],
    })
    editor.commands.setTextSelection(1)
    expect(marked()).toEqual(['teh@1'])
    editor.commands.setTextSelection(10)
    editor.commands.insertContent(' and recieve')
    expect(marked()).toEqual(['teh@1', 'recieve@15'])
  })

  it('marks both halves of a word that is split while it is being typed', () => {
    WRONG.add('wrn').add('og')
    setSpellMarksActive(editor, true)
    editor.commands.setTextSelection(20)
    for (const key of ' wrnog') editor.commands.insertContent(key)
    editor.commands.setTextSelection(24)
    expect(marked()).not.toContain('wrnog@21')
    editor.commands.splitBlock()
    expect(marked()).toEqual(expect.arrayContaining(['wrn@21', 'og@26']))
    WRONG.delete('wrn')
    WRONG.delete('og')
  })

  it('drops an occurrence ignored once and a word ignored everywhere', () => {
    setSpellMarksActive(editor, true)
    ignoreSpellingOnce(editor, { from: 1, to: 4 })
    expect(marked()).toEqual(['teh@13', 'recieve@25'])
    setWordAside('recieve', 'ignored')
    refreshSpellMarks(editor)
    expect(marked()).toEqual(['teh@13'])
  })
})
