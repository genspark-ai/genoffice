import { describe, expect, it } from 'vitest'
import { mapWordControlChars, openLegacyDoc, readLegacyDocText } from '../src/index'
import { buildSyntheticDoc, type SyntheticPiece } from './helpers/synthetic-doc'

describe('synthetic documents: piece table text decoding', () => {
  it('decodes a compressed (Windows-1252) piece: curly quotes, em dash, ellipsis, é', () => {
    const text = '“quoted” — dash… café'
    const doc = openLegacyDoc(buildSyntheticDoc({ pieces: [{ text, compressed: true }] }))
    expect(doc.pieces.pieces).toHaveLength(1)
    expect(doc.pieces.pieces[0]!.compressed).toBe(true)
    expect(doc.fib.ccpText).toBe(text.length)
    expect(doc.rawText).toBe(text)
    expect(doc.text).toBe(text)
    expect(doc.paragraphs).toEqual([text])
  })

  it('decodes an all-ASCII compressed piece', () => {
    const text = 'Plain ASCII only.\rSecond line.'
    const { text: out, paragraphs } = readLegacyDocText(
      buildSyntheticDoc({ pieces: [{ text, compressed: true }] }),
    )
    expect(out).toBe('Plain ASCII only.\nSecond line.')
    expect(paragraphs).toEqual(['Plain ASCII only.', 'Second line.'])
  })

  it('mixes compressed and UTF-16LE pieces and keeps the CP boundaries', () => {
    const pieces: SyntheticPiece[] = [
      { text: 'Hello ', compressed: true },
      { text: '中文 world' },
      { text: '\rend', compressed: true },
    ]
    const doc = openLegacyDoc(buildSyntheticDoc({ pieces }))
    expect(doc.pieces.cps).toEqual([0, 6, 14, 18])
    expect(doc.pieces.pieces.map((piece) => piece.compressed)).toEqual([true, false, true])
    expect(doc.rawText).toBe('Hello 中文 world\rend')
    expect(doc.text).toBe('Hello 中文 world\nend')
    expect(doc.paragraphs).toEqual(['Hello 中文 world', 'end'])
  })

  it('keeps a surrogate pair (emoji) inside a UTF-16LE piece', () => {
    const text = 'Emoji 😀 pair\r'
    const doc = openLegacyDoc(buildSyntheticDoc({ pieces: [{ text }] }))
    // the emoji is two CPs but one code point
    expect(doc.fib.ccpText).toBe(text.length)
    expect(doc.paragraphs).toEqual(['Emoji 😀 pair', ''])
    expect([...doc.paragraphs[0]!].length).toBe(doc.paragraphs[0]!.length - 1)
  })

  it('splits multiple pieces across paragraph marks', () => {
    const pieces: SyntheticPiece[] = [
      { text: 'One\rTwo\r', compressed: true },
      { text: 'Three ' },
      { text: 'four\r', compressed: true },
    ]
    const { paragraphs } = readLegacyDocText(buildSyntheticDoc({ pieces }))
    expect(paragraphs).toEqual(['One', 'Two', 'Three four', ''])
  })

  it('reads only the body story when ccpText is shorter than the piece table total', () => {
    const pieces: SyntheticPiece[] = [
      { text: 'AAA\r', compressed: true },
      { text: 'BBB\r', compressed: true },
      { text: 'CCC\r', compressed: true },
    ]
    const doc = openLegacyDoc(buildSyntheticDoc({ pieces, ccpText: 8 }))
    // the table covers 12 CPs (body + later stories); stage 1 reads [0, ccpText)
    expect(doc.pieces.cps.at(-1)).toBe(12)
    expect(doc.fib.ccpText).toBe(8)
    expect(doc.rawText).toBe('AAA\rBBB\r')
    expect(doc.paragraphs).toEqual(['AAA', 'BBB', ''])
  })

  it('accepts a 0Table document without the fWhichTblStm flag', () => {
    const doc = openLegacyDoc(
      buildSyntheticDoc({ tableStream: '0Table', pieces: [{ text: 'zero table\r' }] }),
    )
    expect(doc.fib.tableStream).toBe('0Table')
    expect(doc.fib.flags & 0x0200).toBe(0)
    expect(doc.text).toBe('zero table\n')
  })

  it('accepts the oldest supported format version (nFib 0x00C1, Word 97)', () => {
    const doc = openLegacyDoc(buildSyntheticDoc({ nFib: 0x00c1, pieces: [{ text: 'word 97\r' }] }))
    expect(doc.fib.nFib).toBe(0x00c1)
    expect(doc.text).toBe('word 97\n')
  })

  it('handles an empty body (ccpText 0)', () => {
    const doc = openLegacyDoc(buildSyntheticDoc({}))
    expect(doc.fib.ccpText).toBe(0)
    expect(doc.pieces.pieces).toEqual([])
    expect(doc.rawText).toBe('')
    expect(doc.text).toBe('')
    expect(doc.paragraphs).toEqual([''])
  })

  it('ignores cfb producer-only streams', () => {
    const doc = openLegacyDoc(buildSyntheticDoc({ pieces: [{ text: 'x\r' }] }))
    // the container the test helper builds carries cfb's 0x01Sh33tJ5 marker
    expect(doc.table.length).toBeGreaterThan(0)
    expect(doc.text).toBe('x\n')
  })
})

describe('Word control-mark mapping', () => {
  it('maps cell marks, line/page breaks, hard/optional hyphens', () => {
    expect(mapWordControlChars('a\u0007b')).toBe('a\tb')
    expect(mapWordControlChars('a\u000bb')).toBe('a\nb')
    expect(mapWordControlChars('a\u000cb')).toBe('a\nb')
    expect(mapWordControlChars('a\u001eb')).toBe('a-b')
    expect(mapWordControlChars('a\u001fb')).toBe('ab')
  })

  it('drops drawing/footnote/annotation anchors and other C0 controls but keeps tabs', () => {
    expect(mapWordControlChars('a\u0001\u0002\u0005\u0008b')).toBe('ab')
    expect(mapWordControlChars('a\u0003\u0004b')).toBe('ab')
    expect(mapWordControlChars('a\tb\nc')).toBe('a\tb\nc')
  })

  it('keeps a field result and drops the instruction', () => {
    expect(mapWordControlChars('Matter \u0013 MERGEFIELD client \u0014Acme Corp\u0015 filed')).toBe(
      'Matter Acme Corp filed',
    )
  })

  it('handles a nested field (result inside a result) and drops an instruction-only field', () => {
    expect(
      mapWordControlChars('\u0013IF \u0013DATE\u00142026\u0015 = "2026" "yes"\u0014yes\u0015'),
    ).toBe('yes')
    expect(mapWordControlChars('a\u0013 PAGE \u0014\u0015b')).toBe('ab')
  })
})
