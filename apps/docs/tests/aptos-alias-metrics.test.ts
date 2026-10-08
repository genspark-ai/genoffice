/**
 * Width contract of the Aptos aliases (fonts.css): Word renders Aptos with the
 * real font; the size-adjusted Carlito faces must reproduce its advances so
 * line breaks match. Truth = Word for Mac probe 2026-09-03, sums of per-char
 * origins at 12pt (pymupdf rawdict), expressed in em.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import * as opentype from 'opentype.js'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/fonts/fonts.css'), 'utf8')
const FONT_DIR = join(__dirname, '../../../packages/ui/src/fonts')
const APP_FONT_DIR = join(__dirname, '../src/renderer/fonts')

interface Face {
  adjust: number
  ranges: Array<[number, number]> | null
  file: string
}

function facesOf(family: string, weight: string): Face[] {
  return [...css.matchAll(/@font-face\s*\{[^}]*\}/g)]
    .map((m) => m[0])
    .filter((f) => f.includes(`font-family: '${family}'`))
    .filter((f) => f.includes(`font-weight: ${weight}`) && f.includes('font-style: normal'))
    .map((f) => {
      const range = /unicode-range:\s*([^;]+);/.exec(f)?.[1]
      return {
        adjust: Number(/size-adjust:\s*([\d.]+)%/.exec(f)![1]) / 100,
        ranges: range
          ? range.split(',').map((r) => {
              const [a, b] = r.trim().replace('U+', '').split('-')
              return [parseInt(a, 16), parseInt(b ?? a, 16)] as [number, number]
            })
          : null,
        file: /Carlito-(\w+)\.ttf/.exec(f)![1],
      }
    })
}

const fonts = new Map<string, opentype.Font>()
function font(file: string): opentype.Font {
  let f = fonts.get(file)
  if (!f) {
    const buf = readFileSync(join(FONT_DIR, `Carlito-${file}.ttf`))
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
  'I confirm that I am the beneficial owner (UBO) of the entity, and the Committee shall review all relevant documentation prior to approval.'

describe('Aptos GO reproduces Word Aptos advances', () => {
  const regular = facesOf('Aptos GO', 'normal')
  const bold = facesOf('Aptos GO', 'bold')

  it('declares general, digit and space faces for both weights', () => {
    for (const faces of [regular, bold]) {
      expect(faces.map((f) => f.ranges?.[0]?.[0] ?? null)).toEqual([
        null,
        0x30,
        0x20,
        0x5f,
        0x21,
        0x7c,
      ])
    }
  })

  it('pangram within 1.5% of Word (24.2718em regular, 25.3275em bold)', () => {
    expect(Math.abs(widthEm(PANGRAM, regular) / 24.2718 - 1)).toBeLessThan(0.015)
    expect(Math.abs(widthEm(PANGRAM, bold) / 25.3275 - 1)).toBeLessThan(0.015)
  })

  it('body sentence within 1.5% of Word (58.2748em)', () => {
    expect(Math.abs(widthEm(BODY, regular) / 58.2748 - 1)).toBeLessThan(0.015)
  })

  it('unscaled Carlito would miss by more than 3%', () => {
    const plain = [{ adjust: 1, ranges: null, file: 'Regular' }]
    expect(widthEm(BODY, plain) / 58.2748).toBeLessThan(0.97)
  })
})

describe('Aptos Display GO reproduces Word Aptos Display advances', () => {
  it('pangram within 2% of Word (22.9488em)', () => {
    expect(
      Math.abs(widthEm(PANGRAM, facesOf('Aptos Display GO', 'normal')) / 22.9488 - 1),
    ).toBeLessThan(0.02)
  })
})

// Word DFonts Aptos hmtx: _ and the en dash 0.46em in every weight (a 91-underscore
// rule at 11.04pt spans 460.4pt in Word), ! / \\ 0.29-0.39em, | 0.27em; Carlito's
// are wider, so the letter scale put a 45-char rule one line over
describe('Aptos GO punctuation classes', () => {
  const regular = facesOf('Aptos GO', 'normal')
  const bold = facesOf('Aptos GO', 'bold')
  const within = (text: string, faces: Face[], em: number, tol = 0.02) =>
    expect(Math.abs(widthEm(text, faces) / em - 1)).toBeLessThan(tol)

  it('underscore rule within 2% of Word (0.46em per glyph, both weights)', () => {
    within('_'.repeat(91), bold, 0.46 * 91)
    within('_'.repeat(91), regular, 0.46 * 91)
    within('\u2013', regular, 0.46)
  })

  it('slashes, bangs and bars follow their own advances', () => {
    within('/', regular, 0.3394)
    within('\\', bold, 0.3867)
    within('!', regular, 0.293)
    within('|', regular, 0.27)
    within('|', bold, 0.27)
  })

  it('letters keep the corpus scale', () => {
    within(BODY, regular, 58.2748, 0.015)
  })
})

// Word's Symbol bullet (DFonts symbol.ttf U+F0B7) inks 0.357em and lays out with
// 1.0054/0.2197 win metrics; the alias scales Liberation Sans' U+25CF (0.4302em),
// keeps the 1.2251em total and puts the baseline 0.36em above the box bottom
// (Chromium's Calibri/Carlito text baseline); overrides apply after size-adjust
describe('Symbol Bullet GO alias', () => {
  const face = [...css.matchAll(/@font-face\s*\{[^}]*\}/g)]
    .map((m) => m[0])
    .find((f) => f.includes("font-family: 'Symbol Bullet GO'"))!
  const adjust = Number(/size-adjust:\s*([\d.]+)%/.exec(face)![1]) / 100

  it('scales the U+25CF ink to the Symbol bullet', () => {
    const buf = readFileSync(join(APP_FONT_DIR, 'LiberationSans-Regular.ttf'))
    const f = opentype.parse(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength))
    const g = f.charToGlyph('\u25cf')
    const ink = ((g.xMax! - g.xMin!) / f.unitsPerEm) * adjust
    expect(Math.abs(ink / 0.3569 - 1)).toBeLessThan(0.01)
  })

  it('carries the Symbol line total at the declared size, baseline on the text baseline', () => {
    const asc = Number(/ascent-override:\s*([\d.]+)%/.exec(face)![1]) / 100
    const desc = Number(/descent-override:\s*([\d.]+)%/.exec(face)![1]) / 100
    expect(Math.abs(((asc + desc) * adjust) / 1.2251 - 1)).toBeLessThan(0.01)
    expect(Math.abs((desc * adjust) / 0.3355 - 1)).toBeLessThan(0.02)
  })
})
