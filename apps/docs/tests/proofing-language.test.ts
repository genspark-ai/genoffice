/** Tools > Language marks the selected text with a language, in the w:lang slot it belongs to. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Run } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { inlineToRuns, runsToInline } from '../src/renderer/editor/convert'
import {
  languageSlot,
  noProofOf,
  proofingLanguageOf,
  setProofingLanguage,
} from '../src/renderer/editor/proofing-language'
import { languageName } from '../src/renderer/components/LanguageDialog'

describe('proofing language', () => {
  let editor: Editor
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
              { type: 'text', text: 'plain ' },
              {
                type: 'text',
                text: 'marked',
                marks: [
                  {
                    type: 'docTextStyle',
                    attrs: { color: 'FF0000', lang: JSON.stringify({ eastAsia: 'ja-JP' }) },
                  },
                ],
              },
            ],
          },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  const langs = () => {
    const out: Array<[string, unknown, unknown]> = []
    editor.state.doc.descendants((node) => {
      if (!node.isText) return
      const mark = node.marks.find((m) => m.type.name === 'docTextStyle')
      out.push([
        node.text!,
        mark?.attrs.lang ? JSON.parse(mark.attrs.lang) : null,
        mark?.attrs.color,
      ])
    })
    return out
  }

  it('files a language under the slot Word uses for it', () => {
    expect(languageSlot('fr-FR')).toBe('val')
    expect(languageSlot('zh-CN')).toBe('eastAsia')
    expect(languageSlot('ko-KR')).toBe('eastAsia')
    expect(languageSlot('he-IL')).toBe('bidi')
  })

  it('a bare caret marks the word it is in', () => {
    editor.commands.setTextSelection(3)
    expect(setProofingLanguage(editor, 'fr-FR')).toBe(true)
    expect(langs().slice(0, 2)).toEqual([
      ['plain', { val: 'fr-FR' }, null],
      [' ', null, undefined],
    ])
    expect(proofingLanguageOf(editor)).toBe('fr-FR')
  })

  it('a caret at the start of a word reads that word, not the run before it', () => {
    editor.commands.setTextSelection({ from: 1, to: 7 })
    setProofingLanguage(editor, 'fr-FR', true)
    editor.commands.setTextSelection(7)
    expect(proofingLanguageOf(editor)).toBe('ja-JP')
    expect(noProofOf(editor)).toBe(false)
  })

  it('a bare caret outside any word marks what is typed next', () => {
    editor.commands.setContent({ type: 'doc', content: [{ type: 'docParagraph' }] })
    editor.commands.setTextSelection(1)
    expect(setProofingLanguage(editor, 'fr-FR', true)).toBe(true)
    expect(proofingLanguageOf(editor)).toBe('fr-FR')
    editor.commands.insertContent('oui')
    expect(langs()).toEqual([['oui', { val: 'fr-FR' }, null]])
    expect(noProofOf(editor)).toBe(true)
  })

  it('reads the slot the text is written in', () => {
    const lang = JSON.stringify({ val: 'en-US', eastAsia: 'zh-CN', bidi: 'he-IL' })
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            {
              type: 'text',
              text: 'word \u4e2d\u6587 \u05e9\u05dc\u05d5\u05dd',
              marks: [{ type: 'docTextStyle', attrs: { lang } }],
            },
          ],
        },
      ],
    })
    editor.commands.setTextSelection({ from: 1, to: 5 })
    expect(proofingLanguageOf(editor)).toBe('en-US')
    editor.commands.setTextSelection({ from: 6, to: 8 })
    expect(proofingLanguageOf(editor)).toBe('zh-CN')
    editor.commands.setTextSelection(7)
    expect(proofingLanguageOf(editor)).toBe('zh-CN')
    editor.commands.setTextSelection({ from: 9, to: 13 })
    expect(proofingLanguageOf(editor)).toBe('he-IL')
  })

  it('marks every run of the selection as that one language, each keeping its formatting', () => {
    editor.commands.setTextSelection({ from: 1, to: 13 })
    expect(setProofingLanguage(editor, 'fr-FR')).toBe(true)
    expect(langs()).toEqual([
      ['plain ', { val: 'fr-FR' }, null],
      ['marked', { val: 'fr-FR' }, 'FF0000'],
    ])
    editor.commands.setTextSelection({ from: 8, to: 12 })
    expect(proofingLanguageOf(editor)).toBe('fr-FR')
    // a language of another script replaces it, and is what reads back
    setProofingLanguage(editor, 'ja-JP')
    expect(proofingLanguageOf(editor)).toBe('ja-JP')
    setProofingLanguage(editor, 'he-IL')
    expect(proofingLanguageOf(editor)).toBe('he-IL')
  })

  it('only touches the selected part of a run', () => {
    editor.commands.setTextSelection({ from: 1, to: 4 })
    setProofingLanguage(editor, 'zh-CN')
    expect(langs().slice(0, 2)).toEqual([
      ['pla', { eastAsia: 'zh-CN' }, null],
      ['in ', null, undefined],
    ])
  })

  it('goes to the run and back unchanged', () => {
    const runs: Run[] = [{ text: 'a', lang: { val: 'en-GB', bidi: 'ar-SA' } }]
    expect(inlineToRuns(runsToInline(runs))).toEqual(runs)
    expect(inlineToRuns(runsToInline([{ text: 'a' }]))).toEqual([{ text: 'a' }])
  })

  it('names a language in the UI language, and falls back to the tag', () => {
    expect(languageName('fr-FR', 'en')).toMatch(/French/)
    expect(languageName('not a tag', 'en')).toBe('not a tag')
  })

  it('marks the selection as not to be checked, and takes the mark off again', () => {
    editor.commands.setTextSelection({ from: 8, to: 12 })
    expect(noProofOf(editor)).toBe(false)
    setProofingLanguage(editor, 'ja-JP', true)
    expect(noProofOf(editor)).toBe(true)
    expect(editor.view.dom.querySelector('span[spellcheck="false"]')?.textContent).toBe('arke')
    // a language picked without touching the box leaves the mark where it is
    setProofingLanguage(editor, 'fr-FR')
    expect(noProofOf(editor)).toBe(true)
    setProofingLanguage(editor, 'fr-FR', false)
    expect(noProofOf(editor)).toBe(false)
    expect(editor.view.dom.querySelector('span[spellcheck="false"]')).toBeNull()
  })

  it('the mark goes to the run and back, on and off', () => {
    const on: Run[] = [{ text: 'a', noProof: true }]
    expect(inlineToRuns(runsToInline(on))).toEqual(on)
    const off: Run[] = [{ text: 'a', noProof: false }]
    expect(inlineToRuns(runsToInline(off))).toEqual(off)
  })
})
