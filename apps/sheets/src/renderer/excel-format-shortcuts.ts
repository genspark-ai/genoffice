/// Excel's Ctrl+Shift+<symbol> number-format and border shortcuts, matched
/// by `event.code` because the shifted digit keys report symbols in `key`.
/// The Home tab's quick buttons share the Excel style codes below.

export const PERCENT_STYLE = '0%'
export const COMMA_STYLE = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'

export function accountingStyle(symbol: string): string {
  const s = `"${symbol}"`
  return `_(${s}* #,##0.00_);_(${s}* \\(#,##0.00\\);_(${s}* "-"??_);_(@_)`
}

export function currencyStyle(symbol: string): string {
  const s = symbol.length === 1 ? symbol : `"${symbol}"`
  return `${s}#,##0.00_);(${s}#,##0.00)`
}

export const RIBBON_CURRENCY_SYMBOL = '$'

export const NUMBER_FORMAT_SHORTCUTS: Readonly<Record<string, string>> = {
  Backquote: 'General',
  Digit1: '#,##0.00',
  Digit2: 'h:mm AM/PM',
  Digit3: 'd-mmm-yy',
  Digit4: currencyStyle(RIBBON_CURRENCY_SYMBOL),
  Digit5: PERCENT_STYLE,
  Digit6: '0.00E+00',
}

const BORDER_SHORTCUTS: Readonly<Record<string, string>> = {
  Digit7: 'outer',
  Minus: 'none',
}

/// Excel for Mac additionally binds outline / no border to Cmd+Opt+0 / Cmd+Opt+-.
const MAC_BORDER_SHORTCUTS: Readonly<Record<string, string>> = {
  Digit0: 'outer',
  Minus: 'none',
}

export interface FormatShortcutKey {
  readonly code: string
  readonly metaKey: boolean
  readonly ctrlKey: boolean
  readonly altKey: boolean
  readonly shiftKey: boolean
}

/// Ribbon command for a keydown, or null when it is not a format shortcut.
export function formatShortcutCommand(event: FormatShortcutKey): string | null {
  if (event.metaKey && event.altKey && !event.ctrlKey && !event.shiftKey) {
    const preset = MAC_BORDER_SHORTCUTS[event.code]
    return preset ? `border:${preset}` : null
  }
  if (!(event.metaKey || event.ctrlKey) || !event.shiftKey || event.altKey) return null
  const pattern = NUMBER_FORMAT_SHORTCUTS[event.code]
  if (pattern) return `format:${pattern}`
  const preset = BORDER_SHORTCUTS[event.code]
  return preset ? `border:${preset}` : null
}
