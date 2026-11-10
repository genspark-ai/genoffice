import { describe, expect, it } from 'vitest'

import {
  CELL_INSET_PX,
  expandAsteriskFill,
  fillRepeatCount,
  fillRewriteForPattern,
  sectionFillToken,
} from '../src/renderer/numfmt-fix'

const NBSP = '\u00a0'

// 1px per code unit keeps the arithmetic readable: room = width - inset.
const measure = (text: string): number => text.length

// Excel's classic four-section accounting format (builtin 44).
const ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'

describe('sectionFillToken', () => {
  it('finds the fill character and its span in a section', () => {
    expect(sectionFillToken('_("$"* #,##0.00_)')).toEqual({ start: 5, end: 7, fill: ' ' })
    expect(sectionFillToken('0*-')).toEqual({ start: 1, end: 3, fill: '-' })
    expect(sectionFillToken('*00')).toEqual({ start: 0, end: 2, fill: '0' })
    expect(sectionFillToken('@*.')).toEqual({ start: 1, end: 3, fill: '.' })
  })

  it('returns null for sections without a fill', () => {
    expect(sectionFillToken('#,##0.00')).toBeNull()
    expect(sectionFillToken('_(@_)')).toBeNull()
    expect(sectionFillToken('')).toBeNull()
  })

  it('ignores asterisks inside quotes, escapes, skip-width tokens and brackets', () => {
    expect(sectionFillToken('"*"0')).toBeNull()
    expect(sectionFillToken('0\\*')).toBeNull()
    expect(sectionFillToken('0_*')).toBeNull()
    expect(sectionFillToken('[$*-409]0')).toBeNull()
    // A trailing bare asterisk has nothing to repeat.
    expect(sectionFillToken('0*')).toBeNull()
  })

  it('honours the last asterisk when a section carries several (ECMA-376 18.8.31)', () => {
    expect(sectionFillToken('0*-*x')).toEqual({ start: 3, end: 5, fill: 'x' })
    expect(sectionFillToken('* "$"* #,##0')).toEqual({ start: 5, end: 7, fill: ' ' })
  })

  it('takes a full astral code point as the fill character', () => {
    const token = sectionFillToken('0*\u{1F600}')
    expect(token?.fill).toBe('\u{1F600}')
    expect(token?.end).toBe(4)
  })
})

describe('fillRewriteForPattern', () => {
  it('marks each section fill with its own sentinel and records the fill characters', () => {
    const rewrite = fillRewriteForPattern(ACCOUNTING)
    expect(rewrite?.fills).toEqual([' ', ' ', ' ', undefined])
    expect(rewrite?.pattern).toBe(
      '_("$""\uE000"#,##0.00_);_("$""\uE001"\\(#,##0.00\\);_("$""\uE002""-"??_);_(@_)',
    )
  })

  it('returns null (cached) for patterns without any fill', () => {
    expect(fillRewriteForPattern('#,##0.00;[Red](#,##0.00)')).toBeNull()
    expect(fillRewriteForPattern('yyyy-mm-dd')).toBeNull()
    expect(fillRewriteForPattern('General')).toBeNull()
  })

  it('keeps the section separators inside quotes intact', () => {
    const rewrite = fillRewriteForPattern('"a;b"* 0')
    expect(rewrite?.pattern).toBe('"a;b""\uE000"0')
  })
})

describe('fillRepeatCount', () => {
  it('fits whole repetitions into the room left beside the text', () => {
    // 100px cell, 5px inset, 20px of text, 7px per fill: floor(75 / 7) = 10.
    expect(fillRepeatCount(100, 20, 7)).toBe(10)
    expect(CELL_INSET_PX).toBe(5)
  })

  it('never goes negative when the text alone overflows', () => {
    expect(fillRepeatCount(30, 40, 7)).toBe(0)
    expect(fillRepeatCount(25, 20, 7)).toBe(0)
  })

  it('returns zero for a fill character without width', () => {
    expect(fillRepeatCount(100, 20, 0)).toBe(0)
    expect(fillRepeatCount(100, 20, Number.NaN)).toBe(0)
  })
})

describe('expandAsteriskFill', () => {
  it('pins the currency symbol left and the amount right in the positive section', () => {
    // Room 20: "<nbsp>$" + fill + "1,234.50<nbsp>" (11 chars) leaves 9 fills.
    expect(expandAsteriskFill(ACCOUNTING, 1234.5, 25, measure)).toBe(
      `${NBSP}$${NBSP.repeat(9)}1,234.50${NBSP}`,
    )
  })

  it('right-anchors the zero dash and keeps a negative in parentheses', () => {
    // Zero section `_("$"* "-"??_)`: "<nbsp>$" + fill + "-<nbsp><nbsp><nbsp>".
    expect(expandAsteriskFill(ACCOUNTING, 0, 25, measure)).toBe(
      `${NBSP}$${NBSP.repeat(14)}-${NBSP}${NBSP}${NBSP}`,
    )
    // Negative section: "<nbsp>$" + fill + "(1,234.50)" — no trailing pad.
    expect(expandAsteriskFill(ACCOUNTING, -1234.5, 25, measure)).toBe(
      `${NBSP}$${NBSP.repeat(8)}(1,234.50)`,
    )
  })

  it('fills the locale-token accounting formats of the prod samples', () => {
    const usd = '_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* "-"??_ ;_-@_ '
    expect(expandAsteriskFill(usd, 18500, 25, measure)).toBe(
      `${NBSP}$${NBSP.repeat(8)}18,500.00${NBSP}`,
    )
    const noSymbol = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
    expect(expandAsteriskFill(noSymbol, 1234.5, 25, measure)).toBe(
      `${NBSP}${NBSP.repeat(10)}1,234.50${NBSP}`,
    )
  })

  it('fills after the number when the token trails it', () => {
    // `0*-` on 5 in a 15px cell: room 10, text "5" → nine dashes.
    expect(expandAsteriskFill('0*-', 5, 15, measure)).toBe(`5${'-'.repeat(9)}`)
    // Leading zeros: `*00` on 42 → room 10 minus "42" → eight zeros.
    expect(expandAsteriskFill('*00', 42, 15, measure)).toBe(`${'0'.repeat(8)}42`)
  })

  it('applies a text-section fill to string values (dot leaders)', () => {
    expect(expandAsteriskFill('@*.', 'Total', 20, measure)).toBe(`Total${'.'.repeat(10)}`)
    // A numeric-only pattern leaves text alone.
    expect(expandAsteriskFill('0*-', 'abc', 20, measure)).toBeNull()
  })

  it('leaves `_x` skip-width tokens as single NBSP pads', () => {
    // `_(` and `_)` each render one NBSP; only the `*` run grows.
    const text = expandAsteriskFill('_(0* _)', 7, 15, measure)
    expect(text).toBe(`${NBSP}7${NBSP.repeat(7)}${NBSP}`)
    expect(text?.startsWith(NBSP + '7')).toBe(true)
    expect(text?.endsWith(NBSP)).toBe(true)
  })

  it('honours only the last asterisk of a section', () => {
    // `0*-*x`: the `*-` is ignored (dropped, as numfmt renders it), `*x` fills.
    expect(expandAsteriskFill('0*-*x', 5, 15, measure)).toBe(`5${'x'.repeat(9)}`)
  })

  it('returns null when nothing fits, leaving the plain text to the normal path', () => {
    // Room 20 vs "<nbsp>$1,234,567.89<nbsp>" (15) fits four fills…
    expect(expandAsteriskFill(ACCOUNTING, 1234567.89, 25, measure)).not.toBeNull()
    // …but room 14 fits none, and a narrower cell still returns null (no ####
    // here: the caller's overflow rule decides).
    expect(expandAsteriskFill(ACCOUNTING, 1234567.89, 19, measure)).toBeNull()
    expect(expandAsteriskFill(ACCOUNTING, 1234567.89, 10, measure)).toBeNull()
  })

  it('returns null for formats without a fill or when the value picks a fill-less section', () => {
    expect(expandAsteriskFill('#,##0.00', 1, 25, measure)).toBeNull()
    expect(expandAsteriskFill('0.00%', 0.5, 25, measure)).toBeNull()
    // Text value under the accounting format lands in `_(@_)`: no fill.
    expect(expandAsteriskFill(ACCOUNTING, 'n/a', 25, measure)).toBeNull()
  })

  it('fills the section a condition selects', () => {
    const conditional = '[>=1000]* #,##0;0.0'
    expect(expandAsteriskFill(conditional, 2500, 15, measure)).toBe(`${NBSP.repeat(5)}2,500`)
    expect(expandAsteriskFill(conditional, 2.5, 15, measure)).toBeNull()
  })

  it('keeps the implicit minus sign ahead of a leading fill', () => {
    // numfmt (like Excel) prefixes the sign to the whole section output.
    expect(expandAsteriskFill('* 0', -5, 15, measure)).toBe(`-${NBSP.repeat(8)}5`)
  })

  it('uses measured widths, not code units, for the repetition count', () => {
    // 7px digits/symbols, 3px fill: room 45 - "$" 7 - "5" 7 = 31 → ten fills.
    const proportional = (text: string): number =>
      Array.from(text).reduce((sum, ch) => sum + (ch === NBSP ? 3 : 7), 0)
    expect(expandAsteriskFill('"$"* 0', 5, 50, proportional)).toBe(`$${NBSP.repeat(10)}5`)
  })
})
