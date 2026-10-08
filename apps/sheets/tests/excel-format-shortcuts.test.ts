import { describe, expect, it } from 'vitest'

import {
  accountingStyle,
  COMMA_STYLE,
  currencyStyle,
  formatShortcutCommand,
  PERCENT_STYLE,
  type FormatShortcutKey,
} from '../src/renderer/excel-format-shortcuts'
import { numfmtOptionsOf, numfmtPreview } from '../src/renderer/numfmt-dialog'

const key = (code: string, extra: Partial<FormatShortcutKey> = {}): FormatShortcutKey => ({
  code,
  metaKey: false,
  ctrlKey: false,
  altKey: false,
  shiftKey: false,
  ...extra,
})
const ctrlShift = (code: string) => key(code, { ctrlKey: true, shiftKey: true })
const cmdShift = (code: string) => key(code, { metaKey: true, shiftKey: true })

describe('formatShortcutCommand', () => {
  it.each([
    ['Backquote', 'format:General'],
    ['Digit1', 'format:#,##0.00'],
    ['Digit2', 'format:h:mm AM/PM'],
    ['Digit3', 'format:d-mmm-yy'],
    ['Digit4', 'format:$#,##0.00_);($#,##0.00)'],
    ['Digit5', 'format:0%'],
    ['Digit6', 'format:0.00E+00'],
    ['Digit7', 'border:outer'],
    ['Minus', 'border:none'],
  ])('Ctrl+Shift+%s and Cmd+Shift+%s map to %s', (code, command) => {
    expect(formatShortcutCommand(ctrlShift(code))).toBe(command)
    expect(formatShortcutCommand(cmdShift(code))).toBe(command)
  })

  it('accepts the Excel for Mac border variants Cmd+Opt+0 and Cmd+Opt+-', () => {
    expect(formatShortcutCommand(key('Digit0', { metaKey: true, altKey: true }), true)).toBe(
      'border:outer',
    )
    expect(formatShortcutCommand(key('Minus', { metaKey: true, altKey: true }), true)).toBe(
      'border:none',
    )
  })

  it('ignores the mac-only border variants on Windows and Linux', () => {
    expect(formatShortcutCommand(key('Digit0', { metaKey: true, altKey: true }))).toBeNull()
    expect(formatShortcutCommand(key('Minus', { metaKey: true, altKey: true }))).toBeNull()
  })

  it('ignores other modifier combinations', () => {
    expect(formatShortcutCommand(key('Digit5', { ctrlKey: true }))).toBeNull()
    expect(formatShortcutCommand(key('Digit5', { shiftKey: true }))).toBeNull()
    expect(
      formatShortcutCommand(key('Digit5', { ctrlKey: true, shiftKey: true, altKey: true })),
    ).toBeNull()
    expect(formatShortcutCommand(key('Digit0', { ctrlKey: true, shiftKey: true }))).toBeNull()
    expect(
      formatShortcutCommand(key('Digit0', { metaKey: true, altKey: true, shiftKey: true })),
    ).toBeNull()
    expect(formatShortcutCommand(ctrlShift('Digit8'))).toBeNull()
  })
})

describe('Home tab quick styles', () => {
  const render = numfmtPreview

  it('use the Excel codes', () => {
    expect(PERCENT_STYLE).toBe('0%')
    expect(COMMA_STYLE).toBe('_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)')
    expect(accountingStyle('$')).toBe(
      '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    )
    expect(currencyStyle('US$')).toBe('"US$"#,##0.00_);("US$"#,##0.00)')
  })

  it('render through numfmt without leaking format tokens', () => {
    const accounting = accountingStyle('$')
    expect(render(accounting, 1234.5)).toBe(' $1,234.50 ')
    expect(render(accounting, -1234.5)).toBe(' $(1,234.50)')
    expect(render(accounting, 0)).toBe(' $-   ')
    expect(render(accounting, 'abc')).toBe(' abc ')
    expect(render(COMMA_STYLE, 1234.5)).toBe(' 1,234.50 ')
    expect(render(COMMA_STYLE, -1234.5)).toBe(' (1,234.50)')
    expect(render(COMMA_STYLE, 0)).toBe(' -   ')
    expect(render(PERCENT_STYLE, 0.125)).toBe('13%')
  })

  it('round-trip into the Format Cells dialog as Accounting', () => {
    expect(numfmtOptionsOf(accountingStyle('$'))).toMatchObject({
      category: 'accounting',
      symbol: '$',
      decimals: 2,
    })
    expect(numfmtOptionsOf(COMMA_STYLE)).toMatchObject({ category: 'accounting', symbol: '' })
  })
})
