import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import type { HfParagraph } from '@genoffice/docx-engine'
import { makeGapHfEl } from '../src/renderer/editor/hf-dom'
import {
  docHyphenationSig,
  loadHyphenation,
  setDocHyphenation,
  type HyphenationConfig,
} from '../src/renderer/editor/hyphenation'
import {
  STATIC_HYPH_CLASS,
  hyphenPieces,
  hyphenateDom,
  unhyphenateDom,
} from '../src/renderer/editor/hyphenation-static'
import { appendNoteRuns } from '../src/renderer/editor/note-dom'

const OFF: HyphenationConfig = {
  on: false,
  dict: null,
  styleDicts: new Map(),
  caps: true,
  limit: 0,
  zonePx: 0,
}
const ON: HyphenationConfig = { ...OFF, on: true, dict: 'en-us' }

const strip = (texts: string[]) =>
  makeGapHfEl({
    kind: 'header',
    value: { text: '', paras: [{ runs: texts.map((text) => ({ text })) } as HfParagraph] },
    pageNo: 1,
    pageTotal: 1,
    columnPx: 300,
  })

const noteRow = (text: string) => {
  const row = document.createElement('div')
  appendNoteRuns(row, { no: 1, text, richParas: [[{ text }]] }, 'footnote')
  return row
}

const textBox = () => {
  const box = document.createElement('div')
  box.className = 'doc-textbox'
  box.innerHTML =
    '<div class="doc-textbox-para"><span>An inter</span><span style="font-weight:700">national</span><span> table</span></div>' +
    '<div class="doc-textbox-para"><span>INTERNATIONAL</span></div>'
  return box
}

describe('hyphenation outside the body editor', () => {
  beforeAll(async () => {
    expect(await loadHyphenation('en-us')).toBe(true)
  })
  afterEach(() => setDocHyphenation(OFF))

  it('leaves every place as it was built in a document that does not hyphenate', () => {
    expect(docHyphenationSig()).toBe('')
    expect(hyphenPieces(['international table'])).toBeNull()
    const header = strip(['international organizational responsibilities'])
    expect(header.className).toBe('page-hf page-hf-header page-gap-hf')
    expect(header.querySelector('.doc-shy')).toBeNull()
    expect(header.querySelector('.page-hf-para span')!.childNodes).toHaveLength(1)
    const note = noteRow('international organizational responsibilities')
    expect(note.className).toBe('')
    expect(note.querySelector('.doc-shy')).toBeNull()
    const box = textBox()
    const before = box.outerHTML
    hyphenateDom(box)
    expect(box.outerHTML).toBe(before)
  })

  it('cuts the texts of a paragraph at the break points, a word across runs as one word', () => {
    setDocHyphenation(ON)
    expect(hyphenPieces(['An international table'])).toEqual([
      ['An in', 'ter', 'na', 'tion', 'al ta', 'ble'],
    ])
    // in-ter-na-tion-al, whichever run a syllable sits in
    expect(hyphenPieces(['An inter', 'national', ' table'])).toEqual([
      ['An in', 'ter'],
      ['', 'na', 'tion', 'al'],
      [' ta', 'ble'],
    ])
    expect(hyphenPieces(['a into the'])).toBeNull()
  })

  it('keeps words in capitals whole when the document says so', () => {
    setDocHyphenation({ ...ON, caps: false })
    expect(hyphenPieces(['INTERNATIONAL'])).toBeNull()
    expect(hyphenPieces(['INTERNATIONAL table'])).toEqual([['INTERNATIONAL ta', 'ble']])
    setDocHyphenation(ON)
    expect(hyphenPieces(['INTERNATIONAL'])).toEqual([['IN', 'TER', 'NA', 'TION', 'AL']])
  })

  it('draws a header strip and a note with the break points and without changing their text', () => {
    setDocHyphenation(ON)
    const header = strip(['An inter', 'national table'])
    expect(header.classList.contains(STATIC_HYPH_CLASS)).toBe(true)
    expect(header.querySelectorAll('.doc-shy')).toHaveLength(5)
    expect(header.querySelector('.page-hf-para')!.textContent).toBe('An international table')
    const note = noteRow('international')
    expect(note.classList.contains(STATIC_HYPH_CLASS)).toBe(true)
    expect(note.querySelectorAll('.doc-shy')).toHaveLength(4)
    expect(note.textContent).toBe('1international')
  })

  it('cuts markup built elsewhere in place and takes the cuts out again', () => {
    setDocHyphenation(ON)
    const box = textBox()
    const before = box.outerHTML
    hyphenateDom(box)
    expect(box.classList.contains(STATIC_HYPH_CLASS)).toBe(true)
    const [first, second] = Array.from(box.querySelectorAll('.doc-textbox-para'))
    expect(first.querySelectorAll('.doc-shy')).toHaveLength(5)
    expect(first.textContent).toBe('An international table')
    // the bold run keeps its syllables and its formatting
    expect(first.children[1].getAttribute('style')).toBe('font-weight:700')
    expect(first.children[1].querySelectorAll('.doc-shy')).toHaveLength(3)
    expect(second.querySelectorAll('.doc-shy')).toHaveLength(4)
    // a second pass adds nothing
    hyphenateDom(box)
    expect(box.querySelectorAll('.doc-shy')).toHaveLength(9)
    unhyphenateDom(box)
    expect(box.outerHTML).toBe(before)
  })

  it('leaves the text of a text box or strip editor to that editor', () => {
    setDocHyphenation(ON)
    const box = document.createElement('div')
    box.className = 'doc-textbox'
    box.innerHTML =
      '<div class="doc-textbox-editor"><div class="doc-textbox-para">international</div></div>'
    hyphenateDom(box)
    expect(box.querySelector('.doc-shy')).toBeNull()
  })

  it('names the rules a strip was cut under', () => {
    setDocHyphenation(ON)
    const on = docHyphenationSig()
    setDocHyphenation({ ...ON, limit: 2 })
    expect(docHyphenationSig()).not.toBe(on)
    setDocHyphenation({ ...ON, caps: false })
    expect(docHyphenationSig()).not.toBe(on)
  })
})
