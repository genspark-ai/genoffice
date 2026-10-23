/**
 * Coverage contract of the 'KR Theme Latin GO' alias (fonts.css): a missing
 * KR face (Noto Sans CJK KR etc.) keeps hangul on Batang while Word draws
 * every non-hangul character with the substituted Latin face. Truth = Word for
 * Mac PDF of prod corpus 20261020 sample 023 (pymupdf rawdict origins): space
 * 0.22em, U+00B7 0.28, curly quotes 0.375, multiplication sign 0.55 in
 * Cambria, never Batang's 0.333 / 0.5 glyphs.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import * as opentype from 'opentype.js'
import { describe, expect, it } from 'vitest'

const FONT_DIR = join(__dirname, '../src/renderer/fonts')
const css = readFileSync(join(FONT_DIR, 'fonts.css'), 'utf8')
const face = [...css.matchAll(/@font-face\s*\{[^}]*\}/g)]
  .map((m) => m[0])
  .find((f) => f.includes("font-family: 'KR Theme Latin GO'"))!
const ranges = /unicode-range:\s*([^;]+);/
  .exec(face)![1]
  .split(',')
  .map((r) => {
    const [a, b] = r.trim().replace('U+', '').split('-')
    return [parseInt(a, 16), parseInt(b ?? a, 16)] as [number, number]
  })
const covers = (ch: string) => {
  const cp = ch.codePointAt(0)!
  return ranges.some(([a, b]) => cp >= a && cp <= b)
}
const buf = readFileSync(join(FONT_DIR, 'Caladea-Regular.ttf'))
const caladea = opentype.parse(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength))
const advEm = (ch: string) => caladea.charToGlyph(ch).advanceWidth! / caladea.unitsPerEm

describe('KR Theme Latin GO alias', () => {
  it('claims the space and Latin punctuation Word draws in Cambria', () => {
    for (const ch of [' ', '0', 'a', '|', '·', '×', '“', '”', '•', '–', '…']) {
      expect(covers(ch), `U+${ch.codePointAt(0)!.toString(16)}`).toBe(true)
    }
  })

  it('leaves hangul, CJK punctuation and fullwidth forms to the KR chain', () => {
    for (const ch of ['가', '한', '　', '「', '，', ' ']) {
      expect(covers(ch), `U+${ch.codePointAt(0)!.toString(16)}`).toBe(false)
    }
  })

  it('reproduces the Word (Cambria) advances', () => {
    expect(advEm(' ')).toBeCloseTo(0.22, 2)
    expect(advEm('·')).toBeCloseTo(0.282, 2)
    expect(advEm('“')).toBeCloseTo(0.375, 2)
    expect(advEm('×')).toBeCloseTo(0.554, 2)
  })
})
