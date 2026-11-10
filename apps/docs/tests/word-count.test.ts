import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  asianCharCount,
  countWords,
  documentTextForWordCount,
  nonAsianWordCount,
} from '../src/renderer/word-count'

describe('word count CJK rule', () => {
  it('counts asian chars one by one, punctuation included', () => {
    expect(asianCharCount('深圳市人民政府')).toBe(7)
    expect(asianCharCount('残疾预防，从我做起。')).toBe(10)
    expect(asianCharCount('（市残联、市卫生健康委）')).toBe(12)
    expect(asianCharCount('ひらがなカタカナ')).toBe(8)
    expect(asianCharCount('한국어')).toBe(3)
  })

  it('counts non-asian words by word, not by character', () => {
    expect(nonAsianWordCount('hello world')).toBe(2)
    expect(nonAsianWordCount("it's a state-of-the-art plan")).toBe(4)
    expect(nonAsianWordCount('2026年8月2日')).toBe(3)
    expect(nonAsianWordCount('深圳市')).toBe(0)
  })

  it('Words = asian chars + non-asian words (Word parity)', () => {
    const text = '按照 WHO 的要求，2025 年底前完成。'
    // 12 asian chars (incl. the fullwidth comma and period); non-asian words "WHO" + "2025" = 2
    expect(asianCharCount(text)).toBe(12)
    expect(nonAsianWordCount(text)).toBe(2)
    expect(countWords(text)).toBe(14)
  })

  it('half-width digits/latin embedded in CJK are words, full-width punctuation is a char', () => {
    expect(countWords('第1章：GenOffice 使用指南')).toBe(9)
    // 7 asian chars (incl. the fullwidth colon) + "1" and "GenOffice" as 2 words
  })

  it('counts space-delimited non-Latin scripts as words', () => {
    expect(nonAsianWordCount('Привет мир')).toBe(2)
    expect(nonAsianWordCount('Γεια σου')).toBe(2)
    expect(nonAsianWordCount('שלום עולם')).toBe(2)
    expect(nonAsianWordCount('مرحبا بالعالم')).toBe(2)
    expect(countWords('Привет мир')).toBe(2)
    expect(countWords('Hello Привет')).toBe(2)
  })

  it('counts Thai, Lao, Tibetan, Myanmar, Georgian, Khmer, Armenian, Devanagari as words', () => {
    expect(nonAsianWordCount('สวัสดี ครับ')).toBe(2)
    expect(nonAsianWordCount('ສະບາຍດີ')).toBe(1)
    expect(nonAsianWordCount('བཀྲ་ཤིས་བདེ་ལེགས་')).toBe(1)
    expect(nonAsianWordCount('မင်္ဂလာပါ')).toBe(1)
    expect(nonAsianWordCount('გამარჯობა მსოფლიო')).toBe(2)
    expect(nonAsianWordCount('សួស្តី ពិភពលោក')).toBe(2)
    expect(nonAsianWordCount('Բարեւ աշխարհ')).toBe(2)
    expect(nonAsianWordCount('नमस्ते दुनिया')).toBe(2)
    expect(countWords('Hello สวัสดี')).toBe(2)
  })

  it('covers the full surrogate pair range for CJK Extensions B-H without counting emoji', () => {
    const extB = String.fromCodePoint(0x20000)
    const extH = String.fromCodePoint(0x3134a)
    expect(asianCharCount(extB)).toBe(1)
    expect(asianCharCount(extH)).toBe(1)
    expect(asianCharCount('😀')).toBe(0)
    expect(asianCharCount('👨‍👩‍👧‍👦')).toBe(0)
  })

  it('does not count supplementary private-use or variation selectors as asian chars', () => {
    expect(asianCharCount(String.fromCodePoint(0xf0000))).toBe(0)
    expect(asianCharCount(String.fromCodePoint(0x100000))).toBe(0)
    expect(asianCharCount(String.fromCodePoint(0xe0100))).toBe(0)
  })
})

describe('document word-count text', () => {
  it('includes readable formula and ruby text without counting note or index markers', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'Alpha ' },
              { type: 'docInlineMath', attrs: { text: 'E = mc 2' } },
              { type: 'text', text: ' beta ' },
              { type: 'docRuby', attrs: { base: 'ruby', rt: 'ルビ' } },
              { type: 'docNoteRef', attrs: { num: 7 } },
              { type: 'docXeMark', attrs: { term: 'index' } },
            ],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'gamma' }] },
        ],
      },
    })

    try {
      expect(editor.state.doc.textContent).not.toContain('E = mc 2')
      const text = documentTextForWordCount(editor.state.doc)
      expect(countWords(text)).toBe(7)
      expect(text).toContain('E = mc 2')
      expect(text).toContain('ruby')
      expect(text).not.toContain('index')
      expect(text).not.toContain('7')
    } finally {
      editor.destroy()
    }
  })

  it('recognizes a paragraph containing only an inline formula as non-empty', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [{ type: 'docInlineMath', attrs: { text: 'x = 1' } }],
          },
        ],
      },
    })

    try {
      const paragraph = editor.state.doc.firstChild!
      expect(paragraph.textContent).toBe('')
      expect(documentTextForWordCount(paragraph)).toBe('x = 1')
    } finally {
      editor.destroy()
    }
  })
})
