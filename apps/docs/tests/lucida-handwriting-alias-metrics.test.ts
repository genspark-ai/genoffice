/**
 * Width contract of the Lucida Handwriting alias (fonts.css): Word renders the
 * Office-bundled script face, the per-case size-adjusted Liberation Sans Italic
 * faces must reproduce its advances so line breaks match. Truth = hmtx sums of
 * Word's Lucida Handwriting Italic (2026-09-24), in em.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import * as opentype from 'opentype.js'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/fonts/fonts.css'), 'utf8')
const FONT_DIR = join(__dirname, '../src/renderer/fonts')

interface Face {
  adjust: number
  ranges: Array<[number, number]> | null
  file: string
}

const faces: Face[] = [...css.matchAll(/@font-face\s*\{[^}]*\}/g)]
  .map((m) => m[0])
  .filter((f) => f.includes("font-family: 'Lucida Handwriting GO'"))
  .map((f) => {
    const range = /unicode-range:\s*([^;]+);/.exec(f)?.[1]
    return {
      adjust: Number(/size-adjust:\s*([\d.]+)%/.exec(f)?.[1] ?? 100) / 100,
      ranges: range
        ? range.split(',').map((r) => {
            const [a, b] = r.trim().replace('U+', '').split('-')
            return [parseInt(a, 16), parseInt(b ?? a, 16)] as [number, number]
          })
        : null,
      file: /LiberationSans-(\w+)\.ttf/.exec(f)![1],
    }
  })

const fonts = new Map<string, opentype.Font>()
function font(file: string): opentype.Font {
  let f = fonts.get(file)
  if (!f) {
    const buf = readFileSync(join(FONT_DIR, `LiberationSans-${file}.ttf`))
    f = opentype.parse(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength))
    fonts.set(file, f)
  }
  return f
}

/** later faces win inside their unicode-range (CSS Fonts 4) */
function widthEm(text: string, list: Face[]): number {
  let em = 0
  for (const ch of text) {
    const cp = ch.codePointAt(0)!
    const face = [...list]
      .reverse()
      .find((f) => !f.ranges || f.ranges.some(([a, b]) => cp >= a && cp <= b))!
    const f = font(face.file)
    em += (f.charToGlyph(ch).advanceWidth! / f.unitsPerEm) * face.adjust
  }
  return em
}

const PANGRAM = 'The quick brown fox jumps over the lazy dog 0123456789'
const CAPS = 'GODLY MARRIAGE EXPERIENCE: KEY TO INNER HEALING'
const BODY =
  'there is no more lovely, charming and fulfilling relationship, and no more powerful testimony, than a godly marriage'
const PROGRAM = '7:30 Worship and Devotional: 12:30 Lunch 2016'

describe('Lucida Handwriting GO reproduces Lucida Handwriting advances', () => {
  it('declares general, lowercase, capital and digit italic faces at normal weight only', () => {
    expect(faces.map((f) => f.ranges?.[0]?.[0] ?? null)).toEqual([null, 0x61, 0x41, 0x30])
    expect(faces.every((f) => f.file === 'Italic')).toBe(true)
  })

  it('pangram within 2.5% (31.9331em)', () => {
    expect(Math.abs(widthEm(PANGRAM, faces) / 31.9331 - 1)).toBeLessThan(0.025)
  })

  it('capital heading within 1% (30.6587em)', () => {
    expect(Math.abs(widthEm(CAPS, faces) / 30.6587 - 1)).toBeLessThan(0.01)
  })

  it('body line within 2.5% (67.4414em), programme line within 1.5% (26.9756em)', () => {
    expect(Math.abs(widthEm(BODY, faces) / 67.4414 - 1)).toBeLessThan(0.025)
    expect(Math.abs(widthEm(PROGRAM, faces) / 26.9756 - 1)).toBeLessThan(0.015)
  })

  it('unscaled Liberation Sans Italic would miss the body line by more than 20%', () => {
    const plain = [{ adjust: 1, ranges: null, file: 'Italic' }]
    expect(widthEm(BODY, plain) / 67.4414).toBeLessThan(0.8)
  })
})
