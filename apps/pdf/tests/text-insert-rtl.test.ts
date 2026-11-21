import { PDFDocument } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { readPdfText } from '../src/main/read-text'
import { applyTextInserts, canDrawText, verifyTextEdits, visualLine } from '../src/main/text-edit'
import type { TextInsertInput } from '../src/shared/ipc'

async function blankPage(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.addPage([300, 400])
  return doc.save({ useObjectStreams: false })
}

function insert(text: string): TextInsertInput {
  return {
    pageIndex: 0,
    origin: [50, 250],
    text,
    fontSize: 18,
    color: [0, 0, 0],
    font: 'arial',
  }
}

/** the word "shalom" with its vowel points, as Hebrew is normally written */
const SHALOM = 'שָׁלוֹם'

describe('visual order of an inserted line', () => {
  it('reverses a pure right-to-left line', () => {
    expect(visualLine('עברית')).toBe('תירבע')
  })

  it('keeps a letter and its vowel points together', () => {
    // reversing codepoints instead of clusters would put the points before the
    // letter they belong to
    const reversed = visualLine(SHALOM)
    expect(reversed.startsWith('ם')).toBe(true)
    expect(reversed.endsWith('שָׁ')).toBe(true)
    expect(visualLine(reversed)).toBe(SHALOM)
  })

  it('swaps paired brackets so they still face the text they enclose', () => {
    expect(visualLine('(\u05e9\u05dc\u05d5\u05dd)')).toBe('(\u05dd\u05d5\u05dc\u05e9)')
    expect(visualLine('\u05d0 [\u05d1]')).toBe('[\u05d1] \u05d0')
  })

  it('leaves directions it cannot lay out alone', () => {
    // mixed directions and digits need a full bidi pass; reordering them here
    // would invent an order neither side asked for
    expect(visualLine('abc עברית')).toBe('abc עברית')
    expect(visualLine('עברית 3')).toBe('עברית 3')
    expect(visualLine('12345')).toBe('12345')
  })
})

describe('inserting a right-to-left line', () => {
  it('draws the word the reader asked for and verifies the save', async () => {
    const text = 'עברית'
    if (!canDrawText(text, 'arial')) return // machine has no Hebrew-capable face
    const result = await applyTextInserts(await blankPage(), [insert(text)])
    expect(result.skipped).toEqual([])
    // pdfium hands a text object's glyphs out in the order they were set, so an
    // unreordered insert extracts mirrored and the save fence refuses the write
    const page = (await readPdfText(result.bytes)).pages[0]!
    expect(page.text).toBe(text)
    expect(await verifyTextEdits(result.bytes, [{ pageIndex: 0, newText: text }])).toEqual([])
  })

  it('verifies the save of a bracketed right-to-left line', async () => {
    const text = '(\u05e9\u05dc\u05d5\u05dd)'
    if (!canDrawText(text, 'arial')) return // machine has no Hebrew-capable face
    const result = await applyTextInserts(await blankPage(), [insert(text)])
    expect(result.skipped).toEqual([])
    expect(await verifyTextEdits(result.bytes, [{ pageIndex: 0, newText: text }])).toEqual([])
  })

  it('draws a line that holds only one bracket of a pair', async () => {
    // the swapped bracket is not in the written line, so the subset must cover it
    const text = '\u05e9\u05dc\u05d5\u05dd)'
    if (!canDrawText(text, 'arial')) return
    const result = await applyTextInserts(await blankPage(), [insert(text)])
    expect(result.skipped).toEqual([])
    expect(await verifyTextEdits(result.bytes, [{ pageIndex: 0, newText: text }])).toEqual([])
  })
})
