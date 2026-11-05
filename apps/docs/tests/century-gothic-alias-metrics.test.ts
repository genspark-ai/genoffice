/**
 * Width contract of the Century Gothic alias (fonts.generated.css): Word renders the
 * Office-bundled font, the advance-matched Liberation Sans derivatives
 * (tools/build-metric-fonts.py) must reproduce its advances so line breaks match.
 * Truth = hmtx sums of Word's Century Gothic / Century Gothic Bold, in em.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import * as opentype from 'opentype.js'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/fonts/fonts.generated.css'), 'utf8')
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
        file: /url\('\.\/([^']+)'\)/.exec(f)![1],
      }
    })
}

const fonts = new Map<string, opentype.Font>()
function font(file: string): opentype.Font {
  let f = fonts.get(file)
  if (!f) {
    const buf = readFileSync(join(FONT_DIR, file))
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
  const cuts = [
    ['normal', 'normal'],
    ['bold', 'normal'],
    ['normal', 'italic'],
    ['bold', 'italic'],
  ] as const

  it('declares one general Liberation Sans face plus one advance-matched derivative per cut', () => {
    for (const [weight, style] of cuts) {
      const faces = facesOf(weight, style)
      expect(faces.filter((f) => !f.ranges)).toHaveLength(1)
      const derived = faces.filter((f) => f.ranges)
      expect(derived).toHaveLength(1)
      expect(derived[0].adjust).toBe(1)
      expect(derived[0].file).toMatch(/^CenturyGothicGO-/)
      const covered = (cp: number) => derived[0].ranges!.some(([a, b]) => cp >= a && cp <= b)
      for (const cp of [0x20, 0x30, 0x43, 0x74, 0xa0, 0xe9, 0x2013, 0x20ac])
        expect(covered(cp)).toBe(true)
    }
  })

  // the bucket faces drew capitals and digits visibly smaller than the lowercase
  it('derived glyphs keep one vertical scale over Liberation Sans', () => {
    const f = font(facesOf('normal', 'normal').find((x) => x.ranges)!.file)
    const c = font('LiberationSans-Regular.ttf')
    const ys = [...'HISTECxsg'].map((ch) => f.charToGlyph(ch).yMax! / c.charToGlyph(ch).yMax!)
    expect(Math.max(...ys) / Math.min(...ys) - 1).toBeLessThan(0.005)
  })

  it('every glyph carries the Word advance: pangram 27.3203em / body 56.3652em regular', () => {
    const faces = facesOf('normal', 'normal')
    expect(widthEm(PANGRAM, faces)).toBeCloseTo(27.3203, 2)
    expect(widthEm(BODY, faces)).toBeCloseTo(56.3652, 2)
    expect(widthEm('CENTURY GOTHIC', faces)).toBeCloseTo(8.5342, 2)
    expect(widthEm('abcdefghijklmnopqrstuvwxyz', faces)).toBeCloseTo(14.0781, 2)
  })

  it('bold: pangram 27.4995em / body 56.1152em', () => {
    const faces = facesOf('bold', 'normal')
    expect(widthEm(PANGRAM, faces)).toBeCloseTo(27.4995, 2)
    expect(widthEm(BODY, faces)).toBeCloseTo(56.1152, 2)
  })

  it('italic shares the upright advances (Word italic hmtx equals regular)', () => {
    expect(widthEm(PANGRAM, facesOf('normal', 'italic'))).toBeCloseTo(27.3203, 2)
  })
})
