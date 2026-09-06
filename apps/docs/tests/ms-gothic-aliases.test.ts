/**
 * Contract of the MS Gothic family aliases in fonts.css: advances match Word's
 * msgothic.ttc (Word probe 2026-09-03) — MS Gothic Latin is a 0.5em monospace,
 * MS PGothic / MS UI Gothic kana are proportional and their brackets half-width.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import * as opentype from 'opentype.js'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const FONTS = join(__dirname, '../src/renderer/fonts')
const css = readFileSync(join(FONTS, 'fonts.css'), 'utf8')

interface Face {
  src: string
  ranges: string
  sizeAdjust: number
}

function faces(family: string): Face[] {
  const out: Face[] = []
  for (const m of css.matchAll(/@font-face\s*\{([^}]*)\}/g)) {
    const body = m[1]
    if (!new RegExp(`font-family:\\s*'${family}'`).test(body)) continue
    out.push({
      src: /src:\s*([^;]+);/.exec(body)?.[1].replace(/\s+/g, ' ') ?? '',
      ranges: /unicode-range:\s*([^;]+);/.exec(body)?.[1].replace(/\s+/g, ' ') ?? '',
      sizeAdjust: Number(/size-adjust:\s*([\d.]+)%/.exec(body)?.[1] ?? 100),
    })
  }
  return out
}

function advanceEmTtf(file: string, ch: string): number {
  const buf = readFileSync(join(FONTS, file))
  const font = opentype.parse(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength))
  return font.charToGlyph(ch).advanceWidth! / font.unitsPerEm
}

describe('MS Gothic GO', () => {
  const all = faces('MS Gothic GO')
  const latin = all.find((f) => claims(f.ranges, 0x61))!

  it('serves Latin from the bundled mono scaled to a 0.5em advance', () => {
    expect(latin.src).toContain('LiberationMono-Regular.ttf')
    expect(latin.ranges).toContain('U+0020-007E')
    for (const ch of ['a', 'W', '0', ' ']) {
      expect(advanceEmTtf('LiberationMono-Regular.ttf', ch) * (latin.sizeAdjust / 100)).toBeCloseTo(
        0.5,
        3,
      )
    }
  })

  it('draws the Latin-1 symbols MS Gothic has fullwidth at a 1em advance', () => {
    expect(claims(latin.ranges, 0xe9)).toBe(true)
    for (const cp of [0xa7, 0xa8, 0xb0, 0xb1, 0xb4, 0xb6, 0xd7, 0xf7]) {
      expect(claims(latin.ranges, cp)).toBe(false)
      const owners = all.filter((f) => claims(f.ranges, cp))
      expect(owners).toHaveLength(1)
      expect(fullwidthAdvance(owners[0], cp)).toBeCloseTo(1, 2)
    }
  })
})

describe.each(['MS PGothic', 'MS UI Gothic'])('%s GO + JA GO', (family) => {
  const latinFaces = faces(`${family} GO`)
  const jaFaces = faces(`${family} JA GO`)
  const all = [...latinFaces, ...jaFaces]
  const at = (cp: number) => all.filter((f) => claims(f.ranges, cp))

  it('draws the fullwidth Latin-1 symbols at 1em like the real face', () => {
    for (const cp of [0xa7, 0xb0, 0xb1, 0xd7, 0xf7]) {
      const owners = at(cp)
      expect(owners).toHaveLength(1)
      expect(fullwidthAdvance(owners[0], cp)).toBeCloseTo(1, 2)
    }
  })

  it('splits Latin and kana ranges across the two families', () => {
    for (const f of latinFaces)
      for (const cp of [0x3042, 0x30a2, 0x30fc, 0x3001, 0x300c])
        expect(claims(f.ranges, cp)).toBe(false)
    for (const f of jaFaces)
      for (const cp of [0x20, 0x61, 0xa7, 0xb0]) expect(claims(f.ranges, cp)).toBe(false)
  })

  it('claims kana, ideographic punctuation and Latin, never kanji', () => {
    expect(at(0x3042)).toHaveLength(1) // hiragana
    expect(at(0x30a2)).toHaveLength(1) // katakana
    expect(at(0x30fc)).toHaveLength(1) // long vowel mark
    expect(at(0x3001)).toHaveLength(1) // ideographic comma
    expect(at(0x300c)).toHaveLength(1) // corner bracket
    expect(at(0x61)).toHaveLength(1)
    expect(at(0x20)).toHaveLength(1)
    expect(at(0x4e00)).toHaveLength(0)
    expect(at(0xff01)).toHaveLength(0) // fullwidth ! stays 1em
  })

  it('scales kana below the 1em Hiragino stand-in and brackets to half width', () => {
    expect(at(0x3042)[0].sizeAdjust).toBeLessThan(100)
    expect(at(0x30a2)[0].sizeAdjust).toBeLessThan(at(0x3042)[0].sizeAdjust)
    expect(at(0x300c)[0].sizeAdjust).toBe(50)
    expect(at(0x3001)[0].sizeAdjust).toBe(66.4)
    for (const f of [at(0x3042)[0], at(0x300c)[0]])
      expect(f.src).toContain("local('Hiragino Sans')")
  })

  it('renders Latin from the bundled Liberation Sans at Word-fitted widths', () => {
    const letters = at(0x61)[0]
    const space = at(0x20)[0]
    expect(letters.src).toContain('LiberationSans-Regular.ttf')
    expect(letters.sizeAdjust).toBeCloseTo(91.37, 2)
    expect(space.sizeAdjust).toBeCloseTo(109.67, 2)
    expect(advanceEmTtf('LiberationSans-Regular.ttf', ' ') * (space.sizeAdjust / 100)).toBeCloseTo(
      0.3047,
      3,
    )
  })
})

it('MS UI Gothic kana run narrower than MS PGothic kana', () => {
  const kana = (family: string) => faces(family).find((f) => claims(f.ranges, 0x3042))!.sizeAdjust
  expect(kana('MS UI Gothic JA GO')).toBeLessThan(kana('MS PGothic JA GO'))
})

/** advance the face yields for a codepoint served from a bundled file */
function fullwidthAdvance(face: Face, cp: number): number {
  const file = /url\('\.\/([^']+)'\)/.exec(face.src)![1]
  const em = file.endsWith('.woff2')
    ? advanceEm(readWoff2(join(FONTS, file)), cp)
    : advanceEmTtf(file, String.fromCodePoint(cp))
  return em * (face.sizeAdjust / 100)
}

function claims(ranges: string, cp: number): boolean {
  return ranges.split(',').some((r) => {
    const m = /U\+([0-9A-F]+)(?:-([0-9A-F]+))?/i.exec(r.trim())
    if (!m) return false
    const lo = parseInt(m[1], 16)
    const hi = m[2] ? parseInt(m[2], 16) : lo
    return cp >= lo && cp <= hi
  })
}
