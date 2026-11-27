/**
 * Out-of-range code-point guards (hardening round 2): pdfium text extraction
 * can surface sentinel unicode values above Unicode's ceiling (0xFFFFFFFF,
 * 0x110000) for malformed/crafted PDFs; String.fromCodePoint throws RangeError
 * on them and the whole conversion dies. The guards skip such chars like other
 * malformed fragments. A crafted PDF cannot reproduce the sentinels on this
 * pdfium build (it clamps broken ToUnicode destinations), so the smallest
 * exported unit fed raw extractor output is tested directly.
 */
import { describe, expect, it } from 'vitest'
import { normalizeRegionalFontArtifacts } from '../src/analyze/chars'
import type { PdfChar } from '../src/ir'
import { mkChar } from './helpers/chars'

/** raw extractor output the way a malformed PDF yields it (P16 D consumers) */
const SENTINEL_CODES = [0xffffffff, 0x110000]

const sentinelChar = (code: number, x: number): PdfChar => {
  const c = mkChar('?', x, { fontFamily: 'Yuanti TC', fontSize: 12 })
  c.code = code // the real extraction never derives text from these
  return c
}

describe('normalizeRegionalFontArtifacts with out-of-range code points', () => {
  const scText =
    '为加强全区各级行政机关政策解读工作规范政策执行强化政策落地保障公众权益办法说明这读们还进关问间'

  it('skips sentinels above the Unicode ceiling instead of throwing', () => {
    const chars = [...scText].map((ch, i) =>
      mkChar(ch, 100 + i * 12, { fontFamily: 'Yuanti TC', fontSize: 12 }),
    )
    chars.unshift(sentinelChar(SENTINEL_CODES[0]!, 60), sentinelChar(SENTINEL_CODES[1]!, 70))
    expect(() => normalizeRegionalFontArtifacts(chars)).not.toThrow()
    // the real simplified evidence still carries the rewrite; the sentinels
    // neither throw nor count as indicator chars
    expect(chars.every((c) => c.fontFamily === 'FangSong')).toBe(true)
  })

  it('leaves families alone when sentinels are all the evidence there is', () => {
    const chars = SENTINEL_CODES.map((code, i) => sentinelChar(code, 100 + i * 12))
    expect(() => normalizeRegionalFontArtifacts(chars)).not.toThrow()
    expect(chars.every((c) => c.fontFamily === 'Yuanti TC')).toBe(true)
  })
})
