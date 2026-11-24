import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  clearIgnoredOnce,
  forgetIgnoredWords,
  ignoreSpellingOnce,
  ignoredOnce,
  misspellingAt,
  setWordAside,
  spellWordAt,
} from '../src/renderer/editor/spell-menu'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'A mispeled word, ' },
            { type: 'text', text: 'gone', marks: [{ type: 'del', attrs: {} }] },
            { type: 'text', text: ' mispeled' },
          ],
        },
      ],
    },
  })
}

const wrong = (word: string) => word === 'mispeled' || word === 'gone'

describe('the misspelled word at a position', () => {
  it('is the word the position is in or touches, wherever the caret sits in it', () => {
    const editor = makeEditor()
    const { doc } = editor.state
    // "A mispeled": the word spans 3..11
    for (const pos of [3, 7, 11])
      expect(spellWordAt(doc, pos)).toEqual({ from: 3, to: 11, word: 'mispeled' })
    expect(spellWordAt(doc, 2)?.word).toBe('A')
    expect(misspellingAt(editor.state, 7, wrong)).toMatchObject({ from: 3, word: 'mispeled' })
    expect(misspellingAt(editor.state, 14, wrong)).toBeNull()
    editor.destroy()
  })

  it('leaves deleted text alone', () => {
    const editor = makeEditor()
    expect(misspellingAt(editor.state, 20, wrong)).toBeNull()
    editor.destroy()
  })

  it('Ignore Once sets one occurrence aside, follows it through edits, and is forgotten on reset', () => {
    const editor = makeEditor()
    ignoreSpellingOnce(editor, { from: 3, to: 11 })
    expect(misspellingAt(editor.state, 7, wrong)).toBeNull()
    // the second occurrence is still reported
    const second = editor.state.doc.content.size - 2
    expect(misspellingAt(editor.state, second, wrong)?.word).toBe('mispeled')
    editor.commands.insertContentAt(1, 'So ')
    expect(ignoredOnce(editor.state, 6, 14)).toBe(true)
    expect(misspellingAt(editor.state, 10, wrong)).toBeNull()
    expect(editor.view.dom.querySelector('[spellcheck="false"]')?.textContent).toBe('mispeled')
    clearIgnoredOnce(editor)
    expect(misspellingAt(editor.state, 10, wrong)?.word).toBe('mispeled')
    editor.destroy()
  })

  it('a word ignored or learned in the session is not reported, whatever the dictionary still says', () => {
    const editor = makeEditor()
    setWordAside('mispeled', 'ignored')
    expect(misspellingAt(editor.state, 7, wrong)).toBeNull()
    // Reset Ignored Words checks it again, and keeps what was added to the dictionary
    setWordAside('gone', 'learned')
    forgetIgnoredWords()
    expect(misspellingAt(editor.state, 7, wrong)?.word).toBe('mispeled')
    setWordAside('mispeled', 'learned')
    forgetIgnoredWords()
    expect(misspellingAt(editor.state, 7, wrong)).toBeNull()
    editor.destroy()
  })
})
