/**
 * Excel-parity fixes for number-format display that Univer's own CELL_CONTENT
 * interceptors get wrong:
 *  - an empty format section renders '' — Univer treats '' as a formatting
 *    failure and falls back to the raw value, so `#,##0;(#,##0);` shows 0
 *  - string cells never enter the formatter, so the 4th (text) section of a
 *    pattern like `0.0_);(0.0);0.0_);@_)` is ignored
 *  - `_x` padding comes out as plain spaces, which layout collapses on
 *    right-aligned cells; re-render with NBSP so accounting columns align
 *  - `*x` fill is dropped; re-render it as the run that fills the cell width
 *  - General is left at stripErrorMargin's 12 significant digits instead of
 *    Excel's "at most 11, shrunk to what the column width can hold"
 *
 * Registered just below Univer's NUMFMT interceptor (priority 10) so its
 * color handling and render cache still run first; this pass only fixes up
 * the resulting value, comparing against the raw cell to see what Univer did.
 */
import {
  CellValueType,
  InterceptorEffectEnum,
  isDefaultFormat,
  numfmt,
  WrapStrategy,
} from '@univerjs/core'
import { ERROR_TYPE_SET, type ErrorType } from '@univerjs/engine-formula'
import { FontCache, getFontStyleString } from '@univerjs/engine-render'
import { ConditionalFormattingService } from '@univerjs/preset-sheets-conditional-formatting'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import { getWorkbookMdw } from './app-constants'
import { isSubstitutedCellFamily } from './cell-font-fallback'
import type { UniverRuntime } from './univer-state'

/// Horizontal room a cell's text never gets: Univer's 2+2px cell padding
/// plus the 1px blur offset. Every width-vs-text rule here (General digit
/// budget, #### overflow, `*x` fill) measures against `width - CELL_INSET_PX`.
export const CELL_INSET_PX = 5

/// Shares the per-workbook max-digit-width with characterWidthToPixels
/// (univer-sync.ts), so a column imported as N chars yields a budget of
/// floor(N) regardless of the render font.
export function generalCharBudget(columnWidthPx: number): number {
  return Math.max(1, Math.floor((columnWidthPx - CELL_INSET_PX) / getWorkbookMdw()))
}

function toScientific(value: number, decimals: number): string {
  const [mantissa, exponent = '+0'] = value.toExponential(decimals).split('e')
  const sign = exponent.startsWith('-') ? '-' : '+'
  return `${mantissa}E${sign}${exponent.replace(/[+-]/, '').padStart(2, '0')}`
}

/**
 * Excel's General display: numfmt already applies the 11-significant-digit
 * cap and the fixed/scientific switch; on top of that, shrink decimals (then
 * the scientific mantissa) until the text fits the column's char budget.
 */
export function formatGeneral(value: number, budget: number): string {
  if (!Number.isFinite(value)) return String(value)
  const base = numfmt.format('General', value)
  if (base.length <= budget) return base
  const abs = Math.abs(value)
  const intLen = (abs < 1 ? 1 : Math.floor(Math.log10(abs)) + 1) + (value < 0 ? 1 : 0)
  for (let dec = Math.min(budget - intLen - 1, 10); dec >= 0; dec -= 1) {
    let t = value.toFixed(dec)
    if (t.includes('.')) t = t.replace(/0+$/, '').replace(/\.$/, '')
    if (Number(t) === 0 && value !== 0) break
    if (t.length <= budget) return t
  }
  for (let dec = 5; dec > 0; dec -= 1) {
    const t = toScientific(value, dec)
    if (t.length <= budget) return t
  }
  return toScientific(value, 0)
}

function safeFormat(pattern: string, value: number | string): string | null {
  try {
    return numfmt.format(pattern, value, { nbsp: true, throws: false })
  } catch {
    return null
  }
}

const formatInfoCache = new Map<string, { type: string; maxDecimals: number; scale: number }>()

/**
 * Excel rounds the value's 15-significant-digit decimal literal half away
 * from zero; numfmt rounds the binary double (`Math.round(v * 10^d)`), so
 * 1.005 → "1.00" and -2.5 under `0` → "-2". Returns the decimal-rounded
 * value when the two disagree, else null. Only single-section fixed-decimal
 * patterns qualify — per-sign sections can carry different decimal counts.
 */
export function decimalRoundForPattern(pattern: string, value: number): number | null {
  if (!Number.isFinite(value) || pattern.includes(';')) return null
  let info = formatInfoCache.get(pattern)
  if (!info) {
    try {
      info = numfmt.getFormatInfo(pattern) as { type: string; maxDecimals: number; scale: number }
    } catch {
      return null
    }
    formatInfoCache.set(pattern, info)
  }
  if (info.type !== 'number' && info.type !== 'grouped' && info.type !== 'percent') return null
  const decimals = info.maxDecimals
  if (!Number.isInteger(decimals) || decimals < 0 || decimals > 14) return null
  const scaled = value * (info.scale || 1)
  const text = scaled.toPrecision(15)
  if (text.includes('e') || text.includes('E')) return null
  const dot = text.indexOf('.')
  if (dot === -1) return null
  const cut = dot + 1 + decimals
  if (cut >= text.length) return null
  const digit = text.charCodeAt(cut) - 48
  if (digit < 0 || digit > 9) return null
  const truncated = Number(text.slice(0, cut).replace(/\.$/, ''))
  const step = (scaled < 0 ? -1 : 1) * 10 ** -decimals
  const rounded = digit >= 5 ? truncated + step : truncated
  return rounded / (info.scale || 1)
}

/// Splits a format pattern on the `;` section separators outside quotes.
function patternSections(pattern: string): string[] {
  const sections: string[] = []
  let current = ''
  let quoted = false
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index] ?? ''
    if (character === '"') quoted = !quoted
    if (character === ';' && !quoted) {
      sections.push(current)
      current = ''
      continue
    }
    if (character === '\\' && !quoted) {
      current += character + (pattern[index + 1] ?? '')
      index += 1
      continue
    }
    current += character
  }
  sections.push(current)
  return sections
}

/**
 * The `*x` fill token a format section honours: its code-unit range in the
 * section and the character to repeat. Quoted literals, `\x` escapes, `_x`
 * skip-width tokens and `[...]` blocks are opaque (an asterisk inside them
 * is not a fill). ECMA-376 18.8.31: a section shall have at most one
 * asterisk, and when it has several, all but the LAST are ignored. A
 * trailing bare `*` is not a fill. Null when the section has no fill.
 */
export function sectionFillToken(
  section: string,
): { start: number; end: number; fill: string } | null {
  let token: { start: number; end: number; fill: string } | null = null
  for (let index = 0; index < section.length; index += 1) {
    const character = section[index]
    if (character === '"') {
      const end = section.indexOf('"', index + 1)
      if (end === -1) return null
      index = end
      continue
    }
    if (character === '[') {
      const end = section.indexOf(']', index + 1)
      if (end === -1) return null
      index = end
      continue
    }
    if (character === '\\' || character === '_') {
      index += 1
      continue
    }
    if (character === '*') {
      const codePoint = section.codePointAt(index + 1)
      if (codePoint === undefined) break
      const fill = String.fromCodePoint(codePoint)
      token = { start: index, end: index + 1 + fill.length, fill }
      index += fill.length
    }
  }
  return token
}

/// Private-use marker per section index: distinct so the rendered text
/// tells which section (hence which fill character) numfmt selected.
function fillSentinel(sectionIndex: number): string {
  return String.fromCharCode(0xe000 + sectionIndex)
}

const FILL_SENTINEL_RANGE = /[\uE000-\uE0FF]/g

interface FillRewrite {
  /// The pattern with each section's honoured `*x` replaced by a quoted
  /// sentinel literal, so numfmt renders the fill position in place.
  pattern: string
  /// Fill character per section index; undefined for sections without one.
  fills: (string | undefined)[]
}

const fillRewriteCache = new Map<string, FillRewrite | null>()

/// Null when no section carries a fill (the common case, cached).
export function fillRewriteForPattern(pattern: string): FillRewrite | null {
  let rewrite = fillRewriteCache.get(pattern)
  if (rewrite !== undefined) return rewrite
  rewrite = null
  if (pattern.includes('*')) {
    const sections = patternSections(pattern)
    const fills: (string | undefined)[] = []
    const rewritten = sections.map((section, index) => {
      const token = sectionFillToken(section)
      fills.push(token?.fill)
      if (!token) return section
      return `${section.slice(0, token.start)}"${fillSentinel(index)}"${section.slice(token.end)}`
    })
    if (fills.some((fill) => fill !== undefined)) {
      rewrite = { pattern: rewritten.join(';'), fills }
    }
  }
  if (fillRewriteCache.size > 5_000) fillRewriteCache.clear()
  fillRewriteCache.set(pattern, rewrite)
  return rewrite
}

/**
 * How many whole fill characters fit beside the rest of the text: Excel
 * repeats `x` until the next copy would overflow the cell, and shows none
 * when the text alone already fills it (the text still renders; a too-wide
 * number then takes the #### rule like any other).
 */
export function fillRepeatCount(
  columnWidthPx: number,
  textWidthPx: number,
  fillWidthPx: number,
): number {
  if (!(fillWidthPx > 0)) return 0
  return Math.max(0, Math.floor((columnWidthPx - CELL_INSET_PX - textWidthPx) / fillWidthPx))
}

/**
 * Excel repeats a pattern's `*x` fill character until the cell is full — the
 * accounting formats use `"$"* ` to pin the currency symbol to the left edge
 * while the amount stays right-aligned; `0*-` draws trailing dashes, `@*.`
 * dot leaders. numfmt drops the token entirely, so re-render the value with
 * the token swapped for a sentinel literal, split the text there and insert
 * the exact run that fills the column. Space fills become NBSP runs so
 * layout keeps them on right-aligned cells (same reasoning as the `_x`
 * padding upgrade). Null when the pattern has no fill, when numfmt picked a
 * section without one, when the marker changed anything but its own
 * position, or when not even one fill character fits.
 */
export function expandAsteriskFill(
  pattern: string,
  value: number | string,
  columnWidthPx: number,
  measure: (text: string) => number,
): string | null {
  const rewrite = fillRewriteForPattern(pattern)
  if (!rewrite) return null
  const marked = safeFormat(rewrite.pattern, value)
  if (marked === null) return null
  const sentinels = marked.match(FILL_SENTINEL_RANGE)
  if (sentinels === null || sentinels.length !== 1) return null
  const sectionIndex = (sentinels[0] as string).charCodeAt(0) - 0xe000
  const rawFill = rewrite.fills[sectionIndex]
  if (rawFill === undefined) return null
  const at = marked.indexOf(sentinels[0] as string)
  const head = marked.slice(0, at)
  const tail = marked.slice(at + 1)
  // The sentinel must be the only difference from the plain render — a
  // literal dropped into a digit run can change grouping or scaling.
  if (head + tail !== safeFormat(pattern, value)) return null
  const fill = rawFill === ' ' ? '\u00a0' : rawFill
  const count = fillRepeatCount(columnWidthPx, measure(head + tail), measure(fill))
  if (count <= 0) return null
  return head + fill.repeat(count) + tail
}

// Only the canonical grouped prefix is allowed before the decimals: a
// trailing comma is Excel's divide-by-1000 scaling, which this rebuild
// does not model.
const DECIMAL_ONLY_PATTERN = /^(?:#,##)?[#0]*\.(0+)%?$/

/**
 * numfmt mangles values whose JS string form is exponential when a fixed-
 * decimal pattern keeps their digits ("0.0000000000" on 1.87e-8 renders
 * "0.87e-800000", 1e-7 renders all zeros). Rebuild the plain decimal text
 * for simple all-zero decimal patterns; anything fancier is left alone.
 */
export function exponentialDecimalRepair(pattern: string, value: number): string | null {
  if (!Number.isFinite(value) || !/[eE]/.test(String(value))) return null
  const match = DECIMAL_ONLY_PATTERN.exec(pattern)
  if (!match) return null
  const decimals = match[1]?.length ?? 0
  if (decimals === 0 || decimals > 20) return null
  const percent = pattern.endsWith('%')
  const scaled = percent ? value * 100 : value
  if (!Number.isFinite(scaled)) return null
  const text = scaled.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
    useGrouping: pattern.includes(','),
  })
  return percent ? `${text}%` : text
}

const NBSP = /\u00a0/g

/// Days between the 1900 and 1904 date-system epochs. numfmt is 1900-only,
/// so 1904 workbooks shift date serials before formatting (display-only; the
/// model keeps the file's original serials so saving stays lossless).
export const DATE_1904_OFFSET = 1462

const datePatternCache = new Map<string, boolean>()

/// Calendar dates only: time-only and elapsed patterns ([h]:mm) render the
/// serial's magnitude, which the epoch shift must not touch.
export function isCalendarDatePattern(pattern: string): boolean {
  let isDate = datePatternCache.get(pattern)
  if (isDate === undefined) {
    try {
      const type = (numfmt.getFormatInfo(pattern) as { type?: string }).type
      isDate = type === 'date' || type === 'datetime'
    } catch {
      isDate = false
    }
    datePatternCache.set(pattern, isDate)
  }
  return isDate
}

const patternTypeCache = new Map<string, string>()

function patternType(pattern: string): string {
  let type = patternTypeCache.get(pattern)
  if (type === undefined) {
    try {
      type = (numfmt.getFormatInfo(pattern) as { type?: string }).type ?? 'unknown'
    } catch {
      type = 'unknown'
    }
    patternTypeCache.set(pattern, type)
  }
  return type
}

/**
 * Excel fills a numeric cell with # when its formatted text is wider than
 * the column (numbers never clip or spill into neighbours). Returns the
 * hash fill, or null when the text fits. Measurement is injected so the
 * rule is testable without a canvas; `scale` calibrates the measurement to
 * Excel's own font metrics (see excelWidthScale) while the fill count stays
 * in render metrics — the hashes are drawn in our font.
 */
export function overflowHashes(
  display: string,
  columnWidthPx: number,
  measure: (text: string) => number,
  scale = 1,
  calibrated = scale !== 1,
): string | null {
  if (display === '') return null
  const available = columnWidthPx - CELL_INSET_PX
  // A calibrated (GDI-normalized) measurement follows Excel's real rule — a
  // digit width of slack before hashing (prod refs pin it in both
  // directions: 36.3px content hashes in a 42px cell, 75.5px shows in 98px).
  // A genuine local face measuring at its GDI width counts as calibrated:
  // its scale clamps to exactly 1 but the numbers are true. Uncalibrated
  // measurements keep a 5% noise band instead: macOS glyph advances run a
  // few percent wide of Excel's GDI metrics (Arial bold "2026/8/30"
  // measures 65.2px here vs ≤63px in Excel), and a borderline spurious ####
  // loses the whole value while a clip loses a few pixels.
  const limit = calibrated ? available - measure('0') * scale : available * 1.05
  if (measure(display) * scale <= limit) return null
  return hashFill(columnWidthPx, measure)
}

/// True for the faces whose GDI digit width is pinned em-exactly: when
/// the genuine face is installed the scale clamps to exactly 1, but the
/// measurement still matches Excel, so the digit-slack rule (not the noise
/// band) applies. Deliberately NOT the wider MDW table — extending Excel's
/// digit-slack semantics to every Calibri cell needs more calibration
/// samples than the one Yu Gothic ref that pinned it.
export function hasGdiDigitCalibration(family: string | undefined): boolean {
  return family !== undefined && SUBSTITUTED_DIGIT_PER_PT[family] !== undefined
}

interface MergedSpanSheet {
  // void: Univer's Nullable<IRange> includes it.
  getMergedCell(
    row: number,
    col: number,
  ): { startColumn: number; endColumn: number } | null | undefined | void
  getColumnWidth(col: number): number
  getColVisible(col: number): boolean
}

/// Display width of a merged cell's column span (hidden columns excluded),
/// null when the cell is not merged. Excel fits General text and decides
/// #### against the whole span, not the anchor column (prod_037: 23566 in a
/// six-column merge whose anchor is 3.1 chars wide rendered 2E+04).
export function mergedSpanWidth(sheet: MergedSpanSheet, row: number, col: number): number | null {
  const merged = sheet.getMergedCell(row, col)
  if (!merged) return null
  let total = 0
  for (let col2 = merged.startColumn; col2 <= merged.endColumn; col2 += 1) {
    if (sheet.getColVisible(col2)) total += sheet.getColumnWidth(col2)
  }
  return total
}

export function hashFill(columnWidthPx: number, measure: (text: string) => number): string | null {
  const hashWidth = measure('#')
  if (!(hashWidth > 0)) return null
  return '#'.repeat(Math.max(1, Math.floor((columnWidthPx - CELL_INSET_PX) / hashWidth)))
}

/// Excel decides overflow with its own GDI metrics. Column pixels follow the
/// Excel MDW (7px/digit for Calibri 11), but when the cell's font is missing
/// locally the canvas measures a substitute (Helvetica ≈9% wider), biasing
/// borderline cells into a spurious ####. Scale such measurements back to
/// the known GDI digit width; installed fonts already measure true.
/// Also the workbook MDW source (measureNormalFontMdw): MDW = digit/pt × pt.
export const EXCEL_DIGIT_PER_PT: Record<string, number> = {
  Calibri: 7 / 11,
  Verdana: 8 / 10,
  // Korean Excel's default; its GDI digit width matches Calibri (MDW 7px at
  // 11pt), while the macOS substitute (Apple SD Gothic Neo, via the
  // styles.css alias) measures digits ~40% wider — Excel-auto-fitted number
  // columns turned into spurious ####.
  'Malgun Gothic': 7 / 11,
  '맑은 고딕': 7 / 11,
  // MDW 8 at 11pt, solved from Excel print geometry of the ja prod refs
  // (five columns match floor((w+16/256)*8) within 1pt; MDW 7 misses by
  // 20%+). The styles.css alias (Carlito) measures Calibri-like 7.3px.
  'Aptos Narrow': 8 / 11,
  // JIS legacy faces keep em/2 digit advances → 8px at 11pt GDI, matching
  // the grid-derived MS PGothic value (ja Excel's classic 72px default
  // column vs Calibri's 64px).
  'ＭＳ Ｐゴシック': 8 / 11,
  'MS PGothic': 8 / 11,
  'ＭＳ ゴシック': 8 / 11,
  'MS Gothic': 8 / 11,
  'ＭＳ Ｐ明朝': 8 / 11,
  'MS PMincho': 8 / 11,
  'ＭＳ 明朝': 8 / 11,
  'MS Mincho': 8 / 11,
  // GB/Big5 legacy faces share the em/2 digit advance.
  宋体: 8 / 11,
  SimSun: 8 / 11,
  NSimSun: 8 / 11,
  新細明體: 8 / 11,
  PMingLiU: 8 / 11,
  細明體: 8 / 11,
  MingLiU: 8 / 11,
}

/// Families whose @font-face alias substitutes a metrically different local
/// face: the family "resolves", but the canvas measures the substitute, so
/// the GDI scale-back must apply even though fontAvailable() is true.
const ALIAS_SUBSTITUTED = new Set(['Malgun Gothic', '맑은 고딕', 'Aptos Narrow'])

/// Substituted faces with em-exact GDI digit widths (ja / Thai from the
/// genuine font files' hmtx; Cordia New from the prod_066 ref print, which
/// pins digits at 4pt per 11pt while the macOS fallback measures ~0.55 em
/// and hashed amounts Excel fits; the UPC spellings are the same designs
/// under the Thai codepage names; px = em × size × 4/3). Every family here
/// has a width-corrected alias in cell-font-fallback.ts, so the canvas
/// digit already lands on this value and the scale sits at ~1 — the entry
/// mainly flags the measurement as calibrated (hasGdiDigitCalibration).
/// Only for the #### scale-back — never for MDW, whose calibration is
/// grid-derived (measureNormalFontMdw) and would round wrong from these
/// exact per-pt values.
const SUBSTITUTED_DIGIT_PER_PT: Record<string, number> = {
  'ＭＳ Ｐゴシック': 0.5 * (4 / 3),
  'MS PGothic': 0.5 * (4 / 3),
  'ＭＳ ゴシック': 0.5 * (4 / 3),
  'MS Gothic': 0.5 * (4 / 3),
  'MS UI Gothic': 0.5 * (4 / 3),
  游ゴシック: 0.556 * (4 / 3),
  游ゴシック体: 0.556 * (4 / 3),
  'Yu Gothic': 0.556 * (4 / 3),
  'Yu Gothic UI': 0.539 * (4 / 3),
  メイリオ: 0.621 * (4 / 3),
  Meiryo: 0.621 * (4 / 3),
  'Meiryo UI': 0.621 * (4 / 3),
  'Cordia New': (4 / 11) * (4 / 3),
  CordiaUPC: (4 / 11) * (4 / 3),
  'Angsana New': 0.33 * (4 / 3),
  AngsanaUPC: 0.33 * (4 / 3),
  'TH SarabunPSK': 0.362 * (4 / 3),
  'TH Sarabun New': 0.362 * (4 / 3),
}

const fontAvailabilityCache = new Map<string, boolean>()

/// Canvas width comparison against generic fallbacks rather than
/// document.fonts.check, whose answer for uninstalled system fonts varies
/// across Chromium versions: a family that renders must differ from at
/// least one generic; a missing one falls back and matches both.
export function fontAvailable(family: string): boolean {
  let known = fontAvailabilityCache.get(family)
  if (known === undefined) {
    known = false
    try {
      const context =
        typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d')
      if (context) {
        const width = (font: string): number => {
          context.font = `16px ${font}`
          return context.measureText('01mWi').width
        }
        known =
          width(`"${family}", monospace`) !== width('monospace') ||
          width(`"${family}", serif`) !== width('serif')
      }
    } catch {
      known = false
    }
    fontAvailabilityCache.set(family, known)
  }
  return known
}

export function excelWidthScale(
  family: string | undefined,
  sizePt: number,
  measureDigit: () => number,
  substituteActive?: boolean,
): number {
  if (!family) return 1
  // The em-exact substituted value wins over the MDW table entry: MDW wants
  // the grid-calibrated 8px/11pt, the hash threshold wants the face's true
  // digit advance (MS PGothic 0.5 em = 7.33px at 11pt).
  const substitutedPerPt = SUBSTITUTED_DIGIT_PER_PT[family]
  const perPt = substitutedPerPt ?? EXCEL_DIGIT_PER_PT[family]
  if (perPt === undefined) return 1
  // fontAvailable is self-polluted for alias-substituted names — our own
  // FontFace registers under the original family — so it only gates
  // families outside both substitution tables.
  if (!ALIAS_SUBSTITUTED.has(family) && substitutedPerPt === undefined && fontAvailable(family)) {
    return 1
  }
  const digit = measureDigit()
  if (!(digit > 0)) return 1
  const scale = (perPt * sizePt) / digit
  // Only our own registered substitute is calibrated in both directions — a
  // narrower substitute (Aptos Narrow -> 96% Carlito) must still trip Excel's
  // #### on GDI overflow. Genuine fonts and uncalibrated fallbacks only ever
  // scale back.
  return (substituteActive ?? isSubstitutedCellFamily(family)) ? scale : Math.min(1, scale)
}

/**
 * Fix-ups for a cell that already went through Univer's NUMFMT interceptor.
 * `displayed` is the value Univer left in the cell, `raw` the model value.
 * Returns the corrected display text, or null to leave the cell alone.
 */
export function fixFormattedValue(
  pattern: string,
  raw: string | number,
  displayed: string | number | boolean | undefined,
): string | null {
  if (typeof raw === 'string') {
    // Excel applies number formats to numbers only, but Univer's NUMFMT
    // coerces numeric-looking text (checkCellValueType lets isRealNum win
    // over t=STRING) — "9853" under a date format rendered as 1926-12-22.
    // Re-derive from the raw text: numfmt applies the @/text section and
    // otherwise returns the text unchanged.
    const text = safeFormat(pattern, raw) ?? raw
    return text !== String(displayed ?? '') ? text : null
  }
  const text = safeFormat(pattern, raw)
  if (text === null) return null
  // Empty section (e.g. `#,##0;(#,##0);` for 0, or `;;;`): Univer fell back
  // to the raw value; Excel renders an empty cell.
  if (text === '') return displayed === '' ? null : ''
  const repaired = exponentialDecimalRepair(pattern, raw)
  if (repaired !== null) {
    return repaired === String(displayed ?? '') ? null : repaired
  }
  // Padding upgrade: only override when the outputs differ by NBSP alone, so
  // any locale-specific rendering Univer did stays untouched.
  // Half-way values: Excel rounds the decimal literal away from zero. Must
  // run before the padding upgrade below — a rounded digit change matters
  // more than NBSP fidelity, and the re-format carries the padding anyway.
  const adjusted = decimalRoundForPattern(pattern, raw)
  if (adjusted !== null && adjusted !== raw) {
    const rounded = safeFormat(pattern, adjusted)
    if (rounded !== null && rounded !== '' && rounded !== text) return rounded
  }
  if (
    typeof displayed === 'string' &&
    text !== displayed &&
    text.replace(NBSP, ' ') === displayed
  ) {
    return text
  }
  return null
}

const JA_FONT_FAMILY =
  /(?:ＭＳ|MS)\s*(?:UI\s+)?[PＰ]?\s*(?:ゴシック|明朝|Gothic|Mincho)|Meiryo|メイリオ|Yu\s?(?:Gothic|Mincho)|游ゴシック|游明朝|Hiragino|ヒラギノ|BIZ\s?UD|Noto\s(?:Sans|Serif)\s(?:JP|CJK\s?JP)/i

const JA_LOCALE_TAG = /\[\$[^\]]*-411\]/

/// `[$\-LCID]` currency token whose symbol is the lone 0x5C code point.
const LEGACY_0X5C_CURRENCY = /\[\$\\-([0-9A-Fa-f]{1,8})\]/

/**
 * JIS legacy currency: ja-authored files store the yen sign at 0x5C
 * (`"\"#,##0` — on Shift-JIS systems U+00A5 occupies the backslash code
 * point) and Excel renders it as U+00A5 with Japanese fonts. Swap the
 * literal in the displayed text when the pattern carries the quoted lone
 * backslash and the context is Japanese (ja font or [$-411] tag), or when
 * a `[$\-LCID]` currency token names a locale whose legacy charset maps
 * 0x5C to a currency sign (Shift-JIS yen, KS X 1001 won).
 */
export function yenLiteralDisplay(
  pattern: string,
  displayed: string,
  fontFamily: string | undefined,
): string | null {
  if (!displayed.includes('\\')) return null
  const token = LEGACY_0X5C_CURRENCY.exec(pattern)
  if (token) {
    // Primary language id of the LCID: 0x11 ja, 0x12 ko.
    const primary = parseInt(token[1] ?? '', 16) & 0x3ff
    if (primary === 0x11) return displayed.replaceAll('\\', '¥')
    if (primary === 0x12) return displayed.replaceAll('\\', '₩')
  }
  if (!pattern.includes('"\\"')) return null
  if (!JA_LOCALE_TAG.test(pattern) && !(fontFamily && JA_FONT_FAMILY.test(fontFamily))) return null
  return displayed.replaceAll('\\', '¥')
}

export function installNumberFormatFix(
  runtime: UniverRuntime,
  isDate1904?: () => boolean,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptorService = injector.get(SheetInterceptorService)
  const cfService = injector.get(ConditionalFormattingService)
  const cache = new Map<string, string | null>()
  // Below the value fixes (9.5) so it sees the final displayed text.
  const yenFix = interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    priority: 9.4,
    effect: InterceptorEffectEnum.Value,
    handler: (cell, location, next) => {
      if (
        !cell ||
        cell.t !== CellValueType.NUMBER ||
        typeof cell.v !== 'string' ||
        !cell.v.includes('\\')
      ) {
        return next(cell)
      }
      const style = location.workbook.getStyles().getStyleByCell(cell)
      const pattern = style?.n?.pattern
      if (typeof pattern !== 'string') return next(cell)
      const swapped = yenLiteralDisplay(pattern, cell.v, style?.ff ?? undefined)
      return swapped === null ? next(cell) : next({ ...cell, v: swapped })
    },
  })
  const valueFix = interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Below NUMFMT (10): Univer formats first (keeping its section colors and
    // render cache); this pass only corrects the value it left behind.
    priority: 9.5,
    effect: InterceptorEffectEnum.Value,
    handler: (cell, location, next) => {
      if (!cell || cell.p != null) return next(cell)
      if (cell.t === CellValueType.BOOLEAN) {
        // Excel hashes a too-wide TRUE/FALSE like any non-text value — a
        // logical never clips (ref prints ### for Arial FALSE in a 32px
        // column). Univer stores the label as 0/1.
        const style = location.workbook.getStyles().getStyleByCell(cell)
        if (
          style?.tb !== WrapStrategy.WRAP &&
          !style?.tr?.a &&
          !style?.tr?.v &&
          !location.worksheet.getMergedCell(location.row, location.col)
        ) {
          const fontString = getFontStyleString(style ?? undefined).fontString
          const measure = (text: string) => FontCache.getMeasureText(text, fontString).width
          const hashes = overflowHashes(
            cell.v === 0 || cell.v === false ? 'FALSE' : 'TRUE',
            location.worksheet.getColumnWidth(location.col),
            measure,
            excelWidthScale(style?.ff ?? undefined, style?.fs ?? 11, () => measure('0')),
            hasGdiDigitCalibration(style?.ff ?? undefined),
          )
          if (hashes !== null) return next({ ...cell, v: hashes, t: CellValueType.NUMBER })
        }
        return next(cell)
      }
      if (cell.t === CellValueType.FORCE_STRING) {
        return next(cell)
      }
      const rawValue = location.rawData?.v
      if (rawValue === undefined || rawValue === null || typeof rawValue === 'boolean') {
        return next(cell)
      }
      // The cached-value fallback (priority 9997) replaces engine error
      // literals with the file's cached value; the raw model still holds the
      // error, so re-deriving the display from it would undo the rescue.
      if (
        typeof rawValue === 'string' &&
        ERROR_TYPE_SET.has(rawValue as ErrorType) &&
        cell.v !== rawValue
      ) {
        return next(cell)
      }
      // Trust the model's type: a numeric string with t=NUMBER is a number
      // (Univer stores some edits that way); anything else stays text.
      const raw =
        typeof rawValue === 'string' &&
        location.rawData?.t === CellValueType.NUMBER &&
        Number.isFinite(Number(rawValue))
          ? Number(rawValue)
          : rawValue
      const style = location.workbook.getStyles().getStyleByCell(cell)
      const pattern = style?.n?.pattern
      // A matched CF rule's dxf number format wins over the cell xf, but the
      // CF interceptor merges styles on the Style effect chain, invisible to
      // this Value pass (and to NUMFMT) — ask the CF service directly and
      // re-format outright when the rule carries a different pattern.
      const cfPattern = cfService.composeStyle(
        location.unitId,
        location.subUnitId,
        location.row,
        location.col,
      )?.style?.n?.pattern
      // NUMFMT ran with the pre-merge style, so its output may reflect the
      // cell xf instead of the rule; comparing the re-formatted text with
      // what it left behind makes this a no-op when they agree.
      // Formula results come from the 1900-based engine (TODAY, DATE, ...),
      // so only file-loaded static serials get the epoch shift.
      const date1904 =
        isDate1904?.() === true &&
        typeof raw === 'number' &&
        location.rawData?.f == null &&
        location.rawData?.si == null
      // Excel's #### rule for numeric cells whose text is wider than the
      // column (a merged cell against its whole visible span). Wrapped/
      // rotated cells and text-formatted values keep Univer's behavior; a
      // negative serial under a date or time format is #### at any width
      // (the 1900 calendar has no negative dates).
      const maybeHash = <T extends typeof cell>(outCell: T, patternUsed: unknown): T => {
        if (typeof raw !== 'number' || typeof patternUsed !== 'string') return outCell
        if (isDefaultFormat(patternUsed)) return outCell
        const type = patternType(patternUsed)
        if (type === 'text' || type === 'unknown') return outCell
        if (style?.tb === WrapStrategy.WRAP || style?.tr?.a || style?.tr?.v) return outCell
        const fontString = getFontStyleString(style ?? undefined).fontString
        const measure = (text: string) => FontCache.getMeasureText(text, fontString).width
        const width =
          mergedSpanWidth(location.worksheet, location.row, location.col) ??
          location.worksheet.getColumnWidth(location.col)
        const negativeDate =
          raw < 0 && !date1904 && (type === 'date' || type === 'datetime' || type === 'time')
        const hashes = negativeDate
          ? hashFill(width, measure)
          : overflowHashes(
              String(outCell.v ?? ''),
              width,
              measure,
              excelWidthScale(style?.ff ?? undefined, style?.fs ?? 11, () => measure('0')),
              hasGdiDigitCalibration(style?.ff ?? undefined),
            )
        if (hashes === null) return outCell
        return { ...outCell, v: hashes, t: CellValueType.NUMBER }
      }
      if (cfPattern !== undefined && !isDefaultFormat(cfPattern)) {
        const cfValue =
          date1904 && isCalendarDatePattern(cfPattern) ? (raw as number) + DATE_1904_OFFSET : raw
        const key = `CF ${cfPattern} ${cfValue}`
        let text = cache.get(key)
        if (text === undefined) {
          text = safeFormat(cfPattern, cfValue)
          if (cache.size > 50_000) cache.clear()
          cache.set(key, text)
        }
        if (text !== null && text !== String(cell.v)) {
          return next(
            maybeHash(
              {
                ...cell,
                v: text,
                t: typeof raw === 'string' ? CellValueType.STRING : CellValueType.NUMBER,
              },
              cfPattern,
            ),
          )
        }
        return next(maybeHash(cell, cfPattern))
      }
      if (isDefaultFormat(pattern)) {
        if (cell.t !== CellValueType.NUMBER || typeof raw !== 'number') return next(cell)
        const budget = generalCharBudget(
          mergedSpanWidth(location.worksheet, location.row, location.col) ??
            location.worksheet.getColumnWidth(location.col),
        )
        const key = `G ${raw} ${budget}`
        let text = cache.get(key)
        if (text === undefined) {
          text = formatGeneral(raw, budget)
          if (cache.size > 50_000) cache.clear()
          cache.set(key, text)
        }
        if (text === null || text === String(cell.v)) return next(cell)
        return next({ ...cell, v: text, t: CellValueType.NUMBER })
      }
      if (date1904 && isCalendarDatePattern(pattern as string)) {
        const key = `1904 ${pattern} ${raw}`
        let text = cache.get(key)
        if (text === undefined) {
          text = safeFormat(pattern as string, (raw as number) + DATE_1904_OFFSET)
          if (cache.size > 50_000) cache.clear()
          cache.set(key, text)
        }
        if (text === null || text === String(cell.v)) return next(maybeHash(cell, pattern))
        return next(maybeHash({ ...cell, v: text, t: CellValueType.NUMBER }, pattern))
      }
      // `*x` fill: expand to the cell width (accounting's left-pinned $; a
      // merged cell fills its whole visible span). Display-only — the model
      // value, formula bar and copies keep the unfilled text. Wrapped and
      // rotated cells keep Univer's render: the fill is a single-line width
      // rule and Excel's own behaviour there is unverified.
      if (
        typeof pattern === 'string' &&
        pattern.includes('*') &&
        (typeof raw === 'number' ||
          (typeof raw === 'string' && !ERROR_TYPE_SET.has(raw as ErrorType))) &&
        style?.tb !== WrapStrategy.WRAP &&
        !style?.tr?.a &&
        !style?.tr?.v
      ) {
        const fontString = getFontStyleString(style ?? undefined).fontString
        const measure = (text: string) => FontCache.getMeasureText(text, fontString).width
        const width =
          mergedSpanWidth(location.worksheet, location.row, location.col) ??
          location.worksheet.getColumnWidth(location.col)
        // Width and font are part of the key: a column resize re-runs the
        // interceptor (the skeleton is rebuilt) and must recompute the run.
        const key = `* ${typeof raw} ${pattern} ${raw} ${width} ${fontString}`
        let expanded = cache.get(key)
        if (expanded === undefined) {
          expanded = expandAsteriskFill(pattern, raw, width, measure)
          if (cache.size > 50_000) cache.clear()
          cache.set(key, expanded)
        }
        if (expanded !== null && expanded !== String(cell.v)) {
          // No hash pass: the fill was sized to the cell, so the value fits
          // by construction (a too-wide value returns null above and takes
          // the normal path, hashes included).
          return next({
            ...cell,
            v: expanded,
            t: typeof raw === 'string' ? CellValueType.STRING : CellValueType.NUMBER,
          })
        }
      }
      const key = `${pattern} ${raw} ${cell.v}`
      let text = cache.get(key)
      if (text === undefined) {
        text = fixFormattedValue(pattern as string, raw, cell.v ?? undefined)
        if (cache.size > 50_000) cache.clear()
        cache.set(key, text)
      }
      if (text === null) return next(maybeHash(cell, pattern))
      return next(
        maybeHash(
          {
            ...cell,
            v: text,
            t: typeof raw === 'string' ? CellValueType.STRING : CellValueType.NUMBER,
          },
          pattern,
        ),
      )
    },
  })
  return {
    dispose() {
      valueFix.dispose()
      yenFix.dispose()
    },
  }
}
