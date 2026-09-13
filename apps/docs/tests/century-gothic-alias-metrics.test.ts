/**
 * Width contract of the Century Gothic alias (fonts.css): Word renders the
 * Office-bundled font, the per-case size-adjusted Liberation Sans faces must
 * reproduce its advances so line breaks match. Truth = hmtx sums of Word's
 * Century Gothic / Century Gothic Bold (2026-09-12), in em.
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

function facesOf(weight: string, style: string): Face[] {
  return [...css.matchAll(/@font-face\s*\{[^}]*\}/g)]
    .map((m) => m[0])
    .filter((f) => f.includes("font-family: 'Century Gothic GO'"))
    .filter((f) => f.includes(`font-weight: ${weight}`) && f.includes(`font-style: ${style}`))
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
}

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
function widthEm(text: string, faces: Face[]): number {
  let em = 0
  for (const ch of text) {
    const cp = ch.codePointAt(0)!
    const face = [...faces]
      .reverse()
      .find((f) => !f.ranges || f.ranges.some(([a, b]) => cp >= a && cp <= b))!
    const f = font(face.file)
    em += (f.charToGlyph(ch).advanceWidth! / f.unitsPerEm) * face.adjust
  }
  return em
}

const PANGRAM = 'The quick brown fox jumps over the lazy dog 0123456789'
const BODY =
  'On the Insert tab, the galleries include items that are designed to coordinate with the overall look of your document.'

describe('Century Gothic GO reproduces Century Gothic advances', () => {
  const regular = facesOf('normal', 'normal')
  const bold = facesOf('bold', 'normal')

  it('declares general, lowercase and capital faces for every weight and style', () => {
    for (const weight of ['normal', 'bold']) {
      for (const style of ['normal', 'italic']) {
        const faces = facesOf(weight, style)
        expect(faces.map((f) => f.ranges?.[0]?.[0] ?? null)).toEqual([null, 0x61, 0x41])
      }
    }
  })

  it('pangram within 1% (27.3203em regular, 27.4995em bold)', () => {
    expect(Math.abs(widthEm(PANGRAM, regular) / 27.3203 - 1)).toBeLessThan(0.01)
    expect(Math.abs(widthEm(PANGRAM, bold) / 27.4995 - 1)).toBeLessThan(0.01)
  })

  it('body sentence within 2% (56.3652em regular, 56.1152em bold)', () => {
    expect(Math.abs(widthEm(BODY, regular) / 56.3652 - 1)).toBeLessThan(0.02)
    expect(Math.abs(widthEm(BODY, bold) / 56.1152 - 1)).toBeLessThan(0.02)
  })

  it('Latin-1 accented letters share the per-case faces (23.9644em regular)', () => {
    const accented =
      "\u00c9l\u00e8ve \u00e0 c\u00f4t\u00e9 de la fen\u00eatre, o\u00f9 l'\u00e9t\u00e9 s'ach\u00e8ve d\u00e9j\u00e0"
    const asciiOnly = regular.map((f) => ({ ...f, ranges: f.ranges && [f.ranges[0]] }))
    expect(Math.abs(widthEm(accented, regular) / 23.9644 - 1)).toBeLessThan(0.035)
    expect(Math.abs(widthEm(accented, asciiOnly) / 23.9644 - 1)).toBeGreaterThan(0.035)
  })

  it('unscaled Liberation Sans would miss the regular body line by more than 5%', () => {
    const plain = [{ adjust: 1, ranges: null, file: 'Regular' }]
    expect(widthEm(BODY, plain) / 56.3652).toBeLessThan(0.95)
  })
})
