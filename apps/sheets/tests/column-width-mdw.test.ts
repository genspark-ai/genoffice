import { WrapStrategy } from '@univerjs/core'
import { afterEach, describe, expect, it } from 'vitest'

import {
  getWorkbookMdw,
  pixelsToCharacterWidth,
  setWorkbookMdw,
} from '../src/renderer/app-constants'
import { generalCharBudget } from '../src/renderer/numfmt-fix'
import {
  characterWidthToPixels,
  measureNormalFontMdw,
  paddedBaseColumnWidth,
  resolveNormalMdwFamily,
  toUniverStyle,
} from '../src/renderer/univer-sync'

afterEach(() => setWorkbookMdw(7))

type Workbook = Parameters<typeof measureNormalFontMdw>[0]

function workbook(overrides: {
  fontFamily?: string
  fontSize?: number
  fontScheme?: 'major' | 'minor'
  normalFontName?: string
  minorEa?: string
}): Workbook {
  const { fontFamily, fontSize, fontScheme, normalFontName, minorEa } = overrides
  return {
    styles: [{ fontFamily, fontSize, fontScheme }],
    ...(normalFontName === undefined ? {} : { normalFontName }),
    ...(minorEa === undefined ? {} : { themeFonts: { major: 'X', minor: 'Y', minorEa } }),
  } as unknown as Workbook
}

describe('workbook MDW', () => {
  it('defaults to Calibri 11 (7px) and lays out in whole points', () => {
    expect(getWorkbookMdw()).toBe(7)
    expect(characterWidthToPixels(10.6640625)).toBeCloseTo((56 * 4) / 3, 9)
  })

  it('widens columns for a Verdana-10 workbook (MDW 8)', () => {
    // DateFormatTests.xlsx col C: width 34.832 chars prints 209pt (278.7px)
    // at MDW 6pt; the hardcoded 7 yielded 244px and wrapped a line early.
    setWorkbookMdw(8)
    expect(characterWidthToPixels(34.83203125)).toBeCloseTo((209 * 4) / 3, 9)
  })

  it('keeps the General digit budget on the same MDW', () => {
    // A Verdana-10 column imported as 40 chars is 240pt = 320px; the 5px
    // inset leaves 39 digits on the same MDW, not (320-5)/7 = 45.
    setWorkbookMdw(8)
    expect(generalCharBudget(characterWidthToPixels(40))).toBe(39)
  })

  it('keeps both conversion directions on the same MDW', () => {
    setWorkbookMdw(8)
    const px = characterWidthToPixels(12)
    expect(Math.abs(pixelsToCharacterWidth(px) - 12)).toBeLessThan(0.05)
  })

  it('clamps nonsense MDW values back to 7', () => {
    setWorkbookMdw(0)
    expect(getWorkbookMdw()).toBe(7)
    setWorkbookMdw(Number.NaN)
    expect(getWorkbookMdw()).toBe(7)
  })

  it('derives the built-in default width from baseColWidth', () => {
    // MDW 7 reproduces the classic 8.7109375 chars Excel writes into files;
    // MDW 8 gives 8.625 chars = 52pt (69.3px), matching live Excel's 8.0-char
    // default column — narrow enough that 10-digit General numbers go scientific.
    setWorkbookMdw(7)
    expect(paddedBaseColumnWidth(null)).toBe(8.7109375)
    setWorkbookMdw(8)
    expect(paddedBaseColumnWidth(null)).toBe(8.625)
    expect(characterWidthToPixels(paddedBaseColumnWidth(null))).toBeCloseTo((52 * 4) / 3, 9)
    expect(generalCharBudget((52 * 4) / 3)).toBe(8)
    expect(paddedBaseColumnWidth(10)).toBe(10.625)
  })
})

describe('measureNormalFontMdw', () => {
  it('derives the ja MDW from the literal cached Normal-font name', () => {
    // prod ja workbooks: <name val="MS PGothic (fullwidth)"/> with
    // scheme="minor" under theme latin Calibri — Excel lays out per
    // MS PGothic (MDW 8, the classic 72px default column), not Calibri 7.
    const file = workbook({
      fontFamily: 'Calibri',
      fontSize: 11,
      fontScheme: 'minor',
      normalFontName: 'ＭＳ Ｐゴシック',
    })
    expect(resolveNormalMdwFamily(file)).toBe('ＭＳ Ｐゴシック')
    expect(measureNormalFontMdw(file)).toBe(8)
  })

  it('strips the ja vertical-text @ prefix', () => {
    const file = workbook({
      fontFamily: 'Calibri',
      fontScheme: 'minor',
      normalFontName: '@ＭＳ ゴシック',
    })
    expect(measureNormalFontMdw(file)).toBe(8)
  })

  it('falls back to the theme minor <a:ea> face when no name is cached', () => {
    const file = workbook({
      fontFamily: 'Calibri',
      fontScheme: 'minor',
      minorEa: 'ＭＳ Ｐゴシック',
    })
    expect(measureNormalFontMdw(file)).toBe(8)
  })

  it('uses the GDI table for Aptos Narrow instead of the Carlito alias', () => {
    // MDW 8 solved from Excel print geometry of the Aptos Narrow prod ref;
    // measuring the styles.css Carlito alias yields Calibri-like 7.
    const file = workbook({ fontFamily: 'Aptos Narrow', fontSize: 11 })
    expect(measureNormalFontMdw(file)).toBe(8)
  })

  it('lays Malgun Gothic 11 workbooks out at the Mac Excel MDW 8', () => {
    // Reference grid (scale 64): 14/11.25/17.125 ch print as 112/89/137 px,
    // floor((w+16/256)*8) exactly; the Windows GDI 7 is 12% short.
    const file = workbook({
      fontFamily: 'Calibri',
      fontSize: 11,
      fontScheme: 'minor',
      normalFontName: '맑은 고딕',
    })
    expect(measureNormalFontMdw(file)).toBe(8)
    expect(measureNormalFontMdw(workbook({ fontFamily: 'Malgun Gothic', fontSize: 11 }))).toBe(8)
  })

  it('keeps other Malgun Gothic sizes on the GDI table', () => {
    const file = workbook({ fontFamily: 'Malgun Gothic', fontSize: 10 })
    expect(measureNormalFontMdw(file)).toBe(6)
  })

  it('keeps the Cordia New hash calibration out of the MDW derivation', () => {
    // The em-exact digit width is #### -only; MDW stays on the canvas
    // fallback (7 in the DOM-less test env), not round(0.485 × 11) = 5.
    const file = workbook({ fontFamily: 'Cordia New', fontSize: 11 })
    expect(measureNormalFontMdw(file)).toBe(7)
  })

  it('keeps the other Thai and Yu Gothic hash calibrations out of the MDW derivation too', () => {
    // These families get width-corrected aliases, so the live canvas digit
    // is the right MDW source by construction; the em-exact hash table must
    // not leak into the grid-derived MDW path (7 in the DOM-less env).
    for (const fontFamily of ['TH SarabunPSK', 'Angsana New', 'Yu Gothic UI']) {
      expect(measureNormalFontMdw(workbook({ fontFamily, fontSize: 11 })), fontFamily).toBe(7)
    }
  })

  it('lays Meiryo workbooks out at MDW 9 under every spelling', () => {
    // Normal font <name val="メイリオ"/> 11pt; Excel's B:H grid
    // fits MDW 9 (0.621em digits → 9.1px), the CJK-name fallback gave 8.
    for (const normalFontName of ['メイリオ', 'Meiryo', 'Meiryo UI']) {
      const file = workbook({
        fontFamily: 'Calibri',
        fontSize: 11,
        fontScheme: 'minor',
        normalFontName,
      })
      expect(resolveNormalMdwFamily(file), normalFontName).toBe(normalFontName)
      expect(measureNormalFontMdw(file), normalFontName).toBe(9)
    }
    expect(measureNormalFontMdw(workbook({ fontFamily: 'メイリオ', fontSize: 11 }))).toBe(9)
    expect(measureNormalFontMdw(workbook({ fontFamily: 'Meiryo', fontSize: 9 }))).toBe(7)
    expect(measureNormalFontMdw(workbook({ fontFamily: 'Meiryo', fontSize: 10 }))).toBe(8)
    expect(measureNormalFontMdw(workbook({ fontFamily: 'Meiryo', fontSize: 12 }))).toBe(10)
  })

  it('scales table entries by the Normal font size', () => {
    expect(measureNormalFontMdw(workbook({ fontFamily: 'Calibri', fontSize: 22 }))).toBe(14)
    expect(measureNormalFontMdw(workbook({ fontFamily: 'Verdana', fontSize: 10 }))).toBe(8)
  })

  it('ignores a differing literal that names no known or renderable face', () => {
    const file = workbook({
      fontFamily: 'Calibri',
      fontSize: 11,
      fontScheme: 'minor',
      normalFontName: 'Nonexistent Face',
    })
    expect(resolveNormalMdwFamily(file)).toBe('Calibri')
    expect(measureNormalFontMdw(file)).toBe(8)
  })

  it('lays Calibri 11 out at the Mac Excel MDW 8, not the GDI 7', () => {
    // Live probes + ref print geometry: 50.86ch → 305pt and 32.44ch → 195pt
    // both fit floor((w+16/256)*8); MDW 7 wraps
    // wide wrap columns a line early and re-fits their rows too tall.
    expect(measureNormalFontMdw(workbook({ fontFamily: 'Calibri', fontSize: 11 }))).toBe(8)
  })

  it('rounds the canvas digit width with the calibrated 0.4 threshold', () => {
    // Arial 10 (7.42px) prints at MDW 8; Arial 11 (8.16px) and Century
    // Gothic 11 (8.13px) print at 8, not 9.
    let measured = 7.42
    const context = {
      set font(_value: string) {},
      measureText: () => ({ width: measured }),
    }
    const documentStub = { createElement: () => ({ getContext: () => context }) }
    Object.defineProperty(globalThis, 'document', { value: documentStub, configurable: true })
    try {
      expect(measureNormalFontMdw(workbook({ fontFamily: 'SomeLatin', fontSize: 10 }))).toBe(8)
      measured = 8.16
      expect(measureNormalFontMdw(workbook({ fontFamily: 'SomeLatin', fontSize: 11 }))).toBe(8)
    } finally {
      Reflect.deleteProperty(globalThis, 'document')
    }
  })

  it('quotes the family for canvas measurement so odd names still measure', () => {
    const fonts: string[] = []
    const context = {
      set font(value: string) {
        fonts.push(value)
      },
      measureText: () => ({ width: 9 }),
    }
    const documentStub = { createElement: () => ({ getContext: () => context }) }
    Object.defineProperty(globalThis, 'document', { value: documentStub, configurable: true })
    try {
      const file = workbook({ fontFamily: '12WeirdDigits', fontSize: 11 })
      expect(measureNormalFontMdw(file)).toBe(9)
      expect(fonts).toContain(`${(11 * 96) / 72}px "12WeirdDigits"`)
    } finally {
      Reflect.deleteProperty(globalThis, 'document')
    }
  })
})

describe('toUniverStyle wrap resolution', () => {
  const base = {
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
  }

  it('emits WRAP for wrapping styles and an explicit OVERFLOW otherwise', () => {
    expect(toUniverStyle({ ...base, wrapText: true } as never).tb).toBe(WrapStrategy.WRAP)
    // A resolved non-wrap cell xf must override a WRAP column style at
    // compose time (sample 60384: col style wraps, A1 explicitly does not).
    expect(toUniverStyle({ ...base, wrapText: false } as never).tb).toBe(WrapStrategy.OVERFLOW)
  })
})
