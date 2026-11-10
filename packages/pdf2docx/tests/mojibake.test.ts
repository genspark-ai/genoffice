/**
 * Mojibake gate: a present-but-wrong ToUnicode map decodes to real code
 * points (no U+FFFD), yet the Latin-1 symbol + accent mix gives it away —
 * such pages ship as bitmaps instead of garbled text.
 */
import { describe, expect, it } from 'vitest'
import { extractPage, looksLikeMojibake, withPdfDocument } from '../src/extract'
import { buildLatin1TextPdf } from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

const codesOf = (s: string) => [...s].map((c) => c.codePointAt(0)!)

// symbol soup the way a Type3 font with a bogus ToUnicode map voices CJK
const GIBBERISH = [
  '¯?vxwn~·Ï¯¿²xvå{¾x{q»2wn{zq»¯¯ÿ»ùüñüóø1o~',
  'µöö½¹ø²Wy»±{Oÿowr»2ÿï~abcwn´»ùüõ³û¿üÿ¸·û²',
  'ù´ö³óÀ~ßýö¿ûó¿ÿ·üówwnÏ¸óóû´·üó (abc def)',
]
// accent-heavy but real: every symbol here is a letter
const PORTUGUESE = [
  'As diferenças na concentração de solutos através das membranas',
  'celulares são geradas e mantidas por mecanismos de transporte que',
  'consomem energia, localizados na membrana celular. É a bomba de Na-K ATPase.',
]

describe('looksLikeMojibake', () => {
  it('flags Latin-1 symbol soup mixed with accented letters', () => {
    expect(looksLikeMojibake(codesOf(GIBBERISH.join('')))).toBe(true)
    // UTF-8 bytes of CJK text read as Latin-1
    const cjk =
      '\u793a\u4f8b\u6587\u672c\u6807\u9898 \u00b7 \u793a\u4f8b\u6587\u672c\u6b63\u6587\u793a\u4f8b\u6587\u672c'
    expect(looksLikeMojibake([...new TextEncoder().encode(cjk)])).toBe(true)
  })

  it('passes real accented text, dot-leader tables of contents and short samples', () => {
    expect(looksLikeMojibake(codesOf(PORTUGUESE.join('')))).toBe(false)
    expect(looksLikeMojibake(codesOf('\u76ee\u6b21' + '·'.repeat(200) + '12'))).toBe(false)
    expect(looksLikeMojibake(codesOf(GIBBERISH[0]!.slice(0, 20)))).toBe(false)
  })
})

describe('mojibake page degrade', () => {
  it('a gibberish page degrades to bad-tounicode, a Portuguese page stays text', async () => {
    const m = await loadPdfium()
    const bad = withPdfDocument(m, await buildLatin1TextPdf(GIBBERISH), (doc: number) =>
      extractPage(m, doc, 0),
    )
    expect(bad.degraded).toBe(true)
    expect(bad.degradedReason).toBe('bad-tounicode')
    const good = withPdfDocument(m, await buildLatin1TextPdf(PORTUGUESE), (doc: number) =>
      extractPage(m, doc, 0),
    )
    expect(good.degraded).toBe(false)
  })
})
