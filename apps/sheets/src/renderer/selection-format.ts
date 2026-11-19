import { BaselineOffset, BooleanNumber, type IStyleData, WrapStrategy } from '@univerjs/core'

import { isAccountingPattern } from './numfmt-dialog'

/// One OOXML indent step rendered as left cell padding, in px — roughly the
/// width of three spaces at the default 11pt font.
/// The padding is the only on-screen model for indent, so converting back
/// (padding / step) must recover the exact step count.
export const INDENT_STEP_PX = 8

/// Univer left padding → OOXML indent steps (0 = none). Non-multiples round
/// to the nearest step so foreign padding still echoes something sane.
export function indentStepsOf(style: IStyleData): number {
  const left = style.pd?.l
  if (typeof left !== 'number' || !Number.isFinite(left) || left <= 0) return 0
  return Math.min(250, Math.max(0, Math.round(left / INDENT_STEP_PX)))
}

/// What the Home ribbon and Format Cells dialog echo back from the active
/// selection.
export interface SelectionFormat {
  readonly fontFamily: string | null
  readonly fontSize: number | null
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  readonly strike: boolean
  readonly vertAlign: 'superscript' | 'subscript' | null
  readonly wrap: boolean
  /// Neutral (wire) alignment names, or null when unset.
  readonly horizontalAlignment: string | null
  readonly verticalAlignment: string | null
  /// OOXML textRotation value (0-180, 255 stacked), or null when unset.
  readonly textRotation: number | null
  /// OOXML indent steps derived from the left padding; 0 = no indent.
  readonly indent: number
  readonly fontColor: string | null
  readonly fillColor: string | null
  readonly numberFormat: string
  /// Hyperlink target of the anchor cell ('#Sheet!A1' internal, URL
  /// external), or null when the cell has no link.
  readonly link: string | null
}

/// Univer colors arrive as #RRGGBB, #RRGGBBAA, or rgb()/rgba() strings; the
/// echo needs the plain 6-digit hex that color inputs accept.
export function normalizeHexColor(color: string | null | undefined): string | null {
  if (!color) return null
  const value = color.trim()
  if (/^#[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(value)) {
    return value.slice(0, 7).toUpperCase()
  }
  const rgb = /^rgba?\(\s*(\d+)\s*[, ]\s*(\d+)\s*[, ]\s*(\d+)/.exec(value)
  if (!rgb) return null
  const hex = rgb
    .slice(1, 4)
    .map((channel) => Math.min(255, Number(channel)).toString(16).padStart(2, '0'))
    .join('')
  return `#${hex}`.toUpperCase()
}

const HORIZONTAL_NAMES: Record<number, string> = {
  1: 'left',
  2: 'center',
  3: 'right',
  4: 'justify',
  6: 'distributed',
}
const VERTICAL_NAMES: Record<number, string> = { 1: 'top', 2: 'center', 3: 'bottom' }

/// Univer text rotation → the OOXML value the wire uses (see edit-journal).
function rotationOf(style: IStyleData): number | null {
  const tr = style.tr
  if (!tr) return null
  if (tr.v === BooleanNumber.TRUE) return 255
  if (typeof tr.a !== 'number' || !Number.isFinite(tr.a) || tr.a === 0) return null
  const angle = Math.max(-90, Math.min(90, Math.round(tr.a)))
  return angle >= 0 ? angle : 90 - angle
}

export function toSelectionFormat(
  style: IStyleData,
  numberFormat: string,
  link: string | null = null,
): SelectionFormat {
  return {
    fontFamily: style.ff ?? null,
    fontSize: style.fs ?? null,
    bold: style.bl === BooleanNumber.TRUE,
    italic: style.it === BooleanNumber.TRUE,
    underline: style.ul?.s === BooleanNumber.TRUE,
    strike: style.st?.s === BooleanNumber.TRUE,
    vertAlign:
      style.va === BaselineOffset.SUPERSCRIPT
        ? 'superscript'
        : style.va === BaselineOffset.SUBSCRIPT
          ? 'subscript'
          : null,
    wrap: style.tb === WrapStrategy.WRAP,
    horizontalAlignment: (style.ht != null && HORIZONTAL_NAMES[style.ht]) || null,
    verticalAlignment: (style.vt != null && VERTICAL_NAMES[style.vt]) || null,
    textRotation: rotationOf(style),
    indent: indentStepsOf(style),
    fontColor: normalizeHexColor(style.cl?.rgb ?? null),
    fillColor: normalizeHexColor(style.bg?.rgb ?? null),
    numberFormat: numberFormat || 'General',
    link,
  }
}

export function selectionFormatEquals(
  a: SelectionFormat | null,
  b: SelectionFormat | null,
): boolean {
  if (a === null || b === null) return a === b
  return (
    a.fontFamily === b.fontFamily &&
    a.fontSize === b.fontSize &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.vertAlign === b.vertAlign &&
    a.wrap === b.wrap &&
    a.horizontalAlignment === b.horizontalAlignment &&
    a.verticalAlignment === b.verticalAlignment &&
    a.textRotation === b.textRotation &&
    a.indent === b.indent &&
    a.fontColor === b.fontColor &&
    a.fillColor === b.fillColor &&
    a.numberFormat === b.numberFormat &&
    a.link === b.link
  )
}

/// Category label for the Number group indicator.
export function numberFormatLabel(pattern: string): string {
  if (!pattern || pattern === 'General') return 'General'
  // Color/condition sections and quoted literals are display-only; category
  // detection must not read date/digit letters inside them.
  const bare = pattern.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')
  if (bare.includes('%')) return 'Percentage'
  if (isAccountingPattern(pattern)) return 'Accounting'
  // Accounting is currency plus the aligning fill/padding tokens (_( and *).
  const quoted = pattern.match(/"[^"]*"/g)?.join('') ?? ''
  const currency = /[$¥€£]/.test(bare + quoted) || /\[\$[^\-\]]/.test(pattern)
  if (currency && bare.includes('_(') && bare.includes('*')) return 'Accounting'
  // Currency symbols appear bare ($#,##0), quoted ("$"#,##0), or as a
  // [$<symbol>-<locale>] prefix — a plain [$-409] locale tag is not one.
  if (currency) return 'Currency'
  if (/E[+-]/i.test(bare)) return 'Scientific'
  if (/\?\/[?\d]/.test(bare)) return 'Fraction'
  if (/[hs]/i.test(bare) && !/[yd]/i.test(bare)) return 'Time'
  if (/[ymdhs]/i.test(bare)) return 'Date'
  if (bare.includes('@')) return 'Text'
  if (/[0#?]/.test(bare)) return 'Number'
  return 'Custom'
}
