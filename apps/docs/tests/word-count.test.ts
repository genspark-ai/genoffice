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
      // Word splits a spaced equation at its spaces, "=" included: Alpha,
      // E, =, mc, 2, beta, ruby, gamma
      expect(countWords(text)).toBe(8)
      expect(text).toContain('E = mc 2')
      expect(text).toContain('ruby')
      expect(text).not.toContain('index')
      expect(text).not.toContain('7')
    } finally {
      editor.destroy()
    }
  })

  it('counts the words of a table nested in a cell, as Word does', () => {
    const inner = { rows: [[{ paras: ['deeper cell'] }]] }
    const model = {
      rows: [
        [{ paras: ['No', 'Nama Kegiatan'] }, { paras: ['Bulan 1'] }],
        [{ paras: ['Akuisisi dataset COG-BCI'], nestedTables: [inner] }, { paras: [''] }],
      ],
    }
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    content: [
                      { type: 'docParagraph', content: [{ type: 'text', text: 'Year one' }] },
                      { type: 'docNestedTable', attrs: { model } },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    })

    try {
      const text = documentTextForWordCount(editor.state.doc)
      // Year one (2) + No, Nama Kegiatan (3) + Bulan 1 (2) + Akuisisi dataset COG-BCI (3) + deeper cell (2)
      expect(countWords(text)).toBe(12)
      expect(text).toContain('Akuisisi dataset')
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

/**
 * Word for Mac 16's own Words statistic for each string (AppleScript `compute
 * statistics`, one string per document, 2026-11-19): the ground truth the
 * non-Asian word rule follows.
 */
const WORD_PROBES: Array<[string, number]> = [
  ['0.75', 1],
  ['p<0.05', 1],
  ['>85%', 1],
  ['<2 detik', 2],
  ['a/b', 1],
  ['x/y/z', 1],
  ['multi-dataset', 1],
  ['real-time', 1],
  ['1,234', 1],
  ['1.000.000', 1],
  ['[1]', 1],
  ['[1] Failory', 2],
  ['(cross-functional)', 1],
  ['e.g.', 1],
  ['https://www.failory.com/blog/startup-failure-rate', 1],
  ['doi: 10.3390/app16010234', 2],
  ['3248\u20133265', 2],
  ['pp. 810\u2013821,', 3],
  ['a \u2013 b', 2],
  ['a - b', 3],
  ['a \u2014 b', 2],
  ['a\u2014b', 2],
  ['\u201cquoted\u201d', 1],
  ['startup\u2019s', 1],
  ["it's", 1],
  ['AI-agentic', 1],
  ['4-6', 1],
  ['2024.\u201d', 1],
  ['X', 1],
  ['X X', 2],
  ['\u2014', 0],
  ['-', 1],
  ['&', 1],
  ['...', 1],
  ['a ... b', 3],
  ['60 Subjek, >30%, p<0.05', 4],
  ['Rp 1.000.000', 2],
  ['v1.2.3', 1],
  ['50%', 1],
  ['#tag', 1],
  ['@user', 1],
  ['[5] M. Attah', 3],
  ['(Tabel 2).', 2],
  ['TBME.2024.3474049', 1],
  ['a\u00a0b', 2],
  ['a\tb', 2],
  ['a\u2026b', 1],
  ['\u2026', 1],
  ['a\u2012b', 1],
  ['a\u2015b', 1],
  ['a\u2212b', 1],
  ['a\u00adb', 1],
  ['a\u200bb', 1],
  ['a\u2010b', 3],
  ['a\u2011b', 1],
  ['a\u2002b', 1],
  ['a\u2009b', 1],
  ['a\u3000b', 2],
  ['\u2022 item', 1],
  ['( a )', 3],
  ['a , b', 3],
  ['a . b', 3],
  ['\u201c a \u201d', 3],
  ['\u2013', 0],
  ['a\u2013', 1],
  ['\u2013a', 1],
  ['a\u2014\u2014b', 2],
  ['1\u20132\u20133', 3],
  ['\u00a7 5', 2],
  ['\u2192', 1],
  ['a\u2192b', 3],
  ['*', 1],
  ['**bold**', 1],
  ['~', 1],
  ['\u00b1', 1],
  ['a\u00b7b', 1],
  ['\u2116 5', 2],
  ['Tahun ke-1', 2],
  ['ke-2', 1],
  ['Rp1.000', 1],
  ['x2', 1],
  ['2x', 1],
  ['5\u00d73', 1],
  ['\u00bd', 1],
  ['\u00b0C', 1],
  ['25 \u00b0C', 2],
  ['a / b', 3],
]

describe('word count matches Word', () => {
  it.each(WORD_PROBES)('%j is %i word(s)', (text, words) => {
    expect(countWords(text)).toBe(words)
  })
})
