import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { nextMisspelling, spellWords } from '../src/renderer/editor/spell-walk'
import { ignoreSpellingOnce, spellingDismissed } from '../src/renderer/editor/spell-menu'

const para = (...content: unknown[]) => ({ type: 'docParagraph', content })
const text = (value: string, marks?: unknown[]) => ({ type: 'text', text: value, marks })

function open(...paragraphs: unknown[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: paragraphs } as never,
  })
}

describe('the words a spelling check walks', () => {
  it('reads a word across a formatting boundary as one, with its place in the document', () => {
    const editor = open(
      para(text('We will rec'), text('ieve', [{ type: 'bold' }]), text(" the dog's toy.")),
    )
    const words = spellWords(editor.state.doc)
    expect(words.map((w) => w.word)).toEqual(['We', 'will', 'recieve', 'the', "dog's", 'toy'])
    const recieve = words[2]
    expect(editor.state.doc.textBetween(recieve.from, recieve.to)).toBe('recieve')
    editor.destroy()
  })

  it('leaves out deleted text, no-proofing text and scripts written without spaces', () => {
    const editor = open(
      para(
        text('keep '),
        text('gonne ', [{ type: 'del', attrs: { id: '1', author: 'a', date: '' } }]),
        text('skipme ', [{ type: 'docTextStyle', attrs: { noProof: true } }]),
        text('\u6f22\u5b57 last'),
      ),
    )
    expect(spellWords(editor.state.doc).map((w) => w.word)).toEqual(['keep', 'last'])
    editor.destroy()
  })

  it('finds the next misspelling inside the range, asking about each word once', () => {
    const editor = open(para(text('teh cat and teh dog')), para(text('teh end')))
    const asked: string[] = []
    const wrong = (word: string) => {
      asked.push(word)
      return word === 'teh'
    }
    const { doc } = editor.state
    const first = nextMisspelling(doc, 0, doc.content.size, wrong)!
    expect(first.from).toBe(1)
    const second = nextMisspelling(doc, first.to, doc.content.size, wrong)!
    expect(doc.textBetween(second.from, second.to)).toBe('teh')
    expect(second.from).toBe(13)
    expect(nextMisspelling(doc, second.to, 20, wrong)).toBeNull()
    expect(asked.filter((w) => w === 'cat')).toHaveLength(1)
    // a caret inside a word starts with that word, and the wrapped half stops short of it
    expect(nextMisspelling(doc, 2, doc.content.size, wrong)!.from).toBe(1)
    expect(nextMisspelling(doc, 0, 2, wrong)).toBeNull()
    editor.destroy()
  })

  it('passes over an occurrence Ignore Once set aside, and stops at the next one', () => {
    const editor = open(para(text('teh cat and teh dog')))
    const { doc } = editor.state
    const wrong = (word: string) => word === 'teh'
    ignoreSpellingOnce(editor, { from: 1, to: 4 })
    const next = nextMisspelling(doc, 0, doc.content.size, wrong, (w) =>
      spellingDismissed(editor.state, w),
    )!
    expect(next.from).toBe(13)
    editor.destroy()
  })
})
