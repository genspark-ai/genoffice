import { describe, expect, it } from 'vitest'
import { openLegacyDoc, readLegacyDocText } from '../src/index'
import {
  FIXTURE_NAMES,
  expectedParagraphs,
  expectedText,
  fixtureBytes,
  type FixtureName,
} from './helpers/fixtures'

/** ccpText of each fixture (from its FIB; also the length of the exported body text) */
const EXPECTED_CCP_TEXT: Record<FixtureName, number> = {
  'text-basic': 163,
  'text-long': 8817,
  'word97-sample': 71,
  formatting: 159,
  styles: 155,
  'page-setup': 16,
  'header-footer': 18,
}

describe('real Word 97-2003 fixtures', () => {
  for (const name of FIXTURE_NAMES) {
    it(`${name}: extracts the body text and paragraphs the .txt export has`, () => {
      const bytes = fixtureBytes(name)
      const doc = openLegacyDoc(bytes)

      expect(doc.fib.nFib).toBe(0x0101)
      expect(doc.fib.tableStream).toBe('1Table')
      expect(doc.fib.encrypted).toBe(false)
      expect(doc.fib.ccpText).toBe(EXPECTED_CCP_TEXT[name])
      expect(doc.fib.ccpFtn).toBe(0)
      expect(doc.fib.ccpEdn).toBe(0)
      expect(doc.fib.lcbClx).toBeGreaterThan(0)
      expect(doc.wordDocument.length).toBeGreaterThan(1024)
      expect(doc.table.length).toBeGreaterThan(doc.fib.lcbClx)

      // one piece per fixture, UTF-16LE, covering the body
      expect(doc.pieces.pieces.length).toBeGreaterThanOrEqual(1)
      expect(doc.pieces.pieces[0]!.compressed).toBe(false)
      expect(doc.pieces.cps[0]).toBe(0)
      expect(doc.pieces.cps.at(-1)).toBeGreaterThanOrEqual(doc.fib.ccpText)

      // CPs count UTF-16 code units, so the raw piece text is exactly ccpText long
      expect(doc.rawText.length).toBe(doc.fib.ccpText)
      expect(doc.text.length).toBe(doc.fib.ccpText)
      expect(doc.text).toBe(expectedText(name))
      expect(doc.paragraphs).toEqual(expectedParagraphs(name))
    })
  }

  it('text-basic keeps the empty paragraph, the tab, the emoji and the Windows-1252 punctuation', () => {
    const { paragraphs } = readLegacyDocText(fixtureBytes('text-basic'))
    expect(paragraphs).toEqual([
      'Legacy DOC text extraction',
      'English paragraph with punctuation: 1, 2, 3.',
      '\u4e2d\u6587\u6bb5\u843d\uff1a\u6df7\u5408\u4e2d\u82f1\u6587 ABC \u6d4b\u8bd5\u3002',
      'Column A\tColumn B',
      '',
      'Emoji and quotes: 😀 “curly” — dash…',
      'Last paragraph.',
      '',
    ])
    // the emoji is a surrogate pair: two CPs, one code point
    expect([...paragraphs[5]!].length).toBeLessThan(paragraphs[5]!.length)
  })

  it('word97-sample ends with the empty paragraph of the final paragraph mark', () => {
    const doc = openLegacyDoc(fixtureBytes('word97-sample'))
    expect(doc.paragraphs).toEqual([
      'Legacy Report',
      'Legacy DOC body text',
      'Second paragraph from Word 97-2003.',
      '',
    ])
    expect(doc.rawText.endsWith('\r')).toBe(true)
  })

  it('text-long has 150 paragraphs', () => {
    const { paragraphs } = readLegacyDocText(fixtureBytes('text-long'))
    expect(paragraphs).toHaveLength(151) // 150 paragraph marks + the final empty one
    expect(paragraphs[0]).toContain('Paragraph 1: The quick brown fox')
    expect(paragraphs[1]).toContain('\u7b2c2\u6bb5')
  })
})
