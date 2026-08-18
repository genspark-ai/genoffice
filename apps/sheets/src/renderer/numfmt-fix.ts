/**
 * Excel-parity fixes for number-format display that Univer's own CELL_CONTENT
 * interceptors get wrong:
 *  - an empty format section renders '' — Univer treats '' as a formatting
 *    failure and falls back to the raw value, so `#,##0;(#,##0);` shows 0
 *  - string cells never enter the formatter, so the 4th (text) section of a
 *    pattern like `0.0_);(0.0);0.0_);@_)` is ignored
 *  - `_x`/`*x` padding comes out as plain spaces, which layout collapses on
 *    right-aligned cells; re-render with NBSP so accounting columns align
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
import { FontCache, getFontStyleString } from '@univerjs/engine-render'
import { ConditionalFormattingService } from '@univerjs/preset-sheets-conditional-formatting'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import { getWorkbookMdw } from './app-constants'
import type { UniverRuntime } from './univer-state'

/// Shares the per-workbook max-digit-width with characterWidthToPixels
/// (univer-sync.ts), so a column imported as N chars yields a budget of
/// floor(N) regardless of the render font.
export function generalCharBudget(columnWidthPx: number): number {
  return Math.max(1, Math.floor((columnWidthPx - 5) / getWorkbookMdw()))
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
 * The literal text a section renders before its `*x` fill: `_x` padding
 * (NBSP), quoted literals, escaped characters, and plain literal symbols.
 * Null when the section has no fill or a placeholder precedes it — the fill
 * must sit in the prefix for the split to be unambiguous.
 */
function literalPrefixBeforeFill(section: string): { prefix: string; fill: string } | null {
  let prefix = ''
  for (let index = 0; index < section.length; index += 1) {
    const character = section[index] ?? ''
    if (character === '*') {
      const fill = section[index + 1]
      if (fill === undefined) return null
      return { prefix, fill: fill === ' ' ? '\u00a0' : fill }
    }
    if (character === '_') {
      prefix += '\u00a0'
      index += 1
      continue
    }
    if (character === '\\') {
      prefix += section[index + 1] ?? ''
      index += 1
      continue
    }
    if (character === '"') {
      const end = section.indexOf('"', index + 1)
      if (end === -1) return null
      prefix += section.slice(index + 1, end)
      index = end
      continue
    }
    // A digit/date placeholder or condition before the fill: unsupported.
    if (/[0#?@[eEdmyhs]/.test(character)) return null
    prefix += character
  }
  return null
}

/**
 * Excel repeats a pattern's `*x` fill character until the cell is full — the
 * accounting formats use `"$"* ` to pin the currency symbol to the left edge
 * while the amount stays right-aligned. numfmt drops the fill entirely, so
 * locate the fill point from the matched section's literal prefix and insert
 * the exact run that fills the column. NBSP fill so layout keeps the run on
 * right-aligned cells (same reasoning as the `_x` padding upgrade).
 */
export function expandAsteriskFill(
  pattern: string,
  value: number,
  columnWidthPx: number,
  measure: (text: string) => number,
): string | null {
  const sections = patternSections(pattern)
  const section =
    value < 0 && sections.length > 1
      ? sections[1]
      : value === 0 && sections.length > 2
        ? sections[2]
        : sections[0]
  if (section === undefined || !section.includes('*')) return null
  const split = literalPrefixBeforeFill(section)
  if (!split) return null
  const text = safeFormat(pattern, value)
  if (text === null || !text.startsWith(split.prefix)) return null
  const unit = measure(split.fill)
  if (!(unit > 0)) return null
  const count = Math.floor((columnWidthPx - 5 - measure(text)) / unit)
  if (count <= 0) return null
  return split.prefix + split.fill.repeat(count) + text.slice(split.prefix.length)
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
): string | null {
  // 2+2 cell padding plus 1px blur offset, matching generalCharBudget.
  const available = columnWidthPx - 5
  if (display === '' || measure(display) * scale <= available) return null
  return hashFill(columnWidthPx, measure)
}

export function hashFill(columnWidthPx: number, measure: (text: string) => number): string | null {
  const hashWidth = measure('#')
  if (!(hashWidth > 0)) return null
  return '#'.repeat(Math.max(1, Math.floor((columnWidthPx - 5) / hashWidth)))
}

/// Excel decides overflow with its own GDI metrics. Column pixels follow the
/// Excel MDW (7px/digit for Calibri 11), but when the cell's font is missing
/// locally the canvas measures a substitute (Helvetica ≈9% wider), biasing
/// borderline cells into a spurious ####. Scale such measurements back to
/// the known GDI digit width; installed fonts already measure true.
const EXCEL_DIGIT_PER_PT: Record<string, number> = { Calibri: 7 / 11, Verdana: 8 / 10 }

const fontAvailabilityCache = new Map<string, boolean>()

/// Canvas width comparison against generic fallbacks rather than
/// document.fonts.check, whose answer for uninstalled system fonts varies
/// across Chromium versions: a family that renders must differ from at
/// least one generic; a missing one falls back and matches both.
function fontAvailable(family: string): boolean {
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
): number {
  if (!family) return 1
  const perPt = EXCEL_DIGIT_PER_PT[family]
  if (perPt === undefined || fontAvailable(family)) return 1
  const digit = measureDigit()
  if (!(digit > 0)) return 1
  return Math.min(1, (perPt * sizePt) / digit)
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
    // Univer never formats plain-text cells: apply the text (4th) section.
    if (displayed !== raw) return null
    const text = safeFormat(pattern, raw)
    return text !== null && text !== raw ? text : null
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

export function installNumberFormatFix(
  runtime: UniverRuntime,
  isDate1904?: () => boolean,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptorService = injector.get(SheetInterceptorService)
  const cfService = injector.get(ConditionalFormattingService)
  const cache = new Map<string, string | null>()
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Below NUMFMT (10): Univer formats first (keeping its section colors and
    // render cache); this pass only corrects the value it left behind.
    priority: 9.5,
    effect: InterceptorEffectEnum.Value,
    handler: (cell, location, next) => {
      if (!cell || cell.p != null) return next(cell)
      if (cell.t === CellValueType.BOOLEAN || cell.t === CellValueType.FORCE_STRING) {
        return next(cell)
      }
      const raw = location.rawData?.v
      if (raw === undefined || raw === null || typeof raw === 'boolean') return next(cell)
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
      // column. Wrapped/rotated/merged cells and text-formatted values keep
      // Univer's behavior; a negative serial under a date or time format is
      // #### at any width (the 1900 calendar has no negative dates).
      const maybeHash = <T extends typeof cell>(outCell: T, patternUsed: unknown): T => {
        if (typeof raw !== 'number' || typeof patternUsed !== 'string') return outCell
        if (isDefaultFormat(patternUsed)) return outCell
        const type = patternType(patternUsed)
        if (type === 'text' || type === 'unknown') return outCell
        if (style?.tb === WrapStrategy.WRAP || style?.tr?.a || style?.tr?.v) return outCell
        if (location.worksheet.getMergedCell(location.row, location.col)) return outCell
        const fontString = getFontStyleString(style ?? undefined).fontString
        const measure = (text: string) => FontCache.getMeasureText(text, fontString).width
        const width = location.worksheet.getColumnWidth(location.col)
        const negativeDate =
          raw < 0 && !date1904 && (type === 'date' || type === 'datetime' || type === 'time')
        const hashes = negativeDate
          ? hashFill(width, measure)
          : overflowHashes(
              String(outCell.v ?? ''),
              width,
              measure,
              excelWidthScale(style?.ff ?? undefined, style?.fs ?? 11, () => measure('0')),
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
        const budget = generalCharBudget(location.worksheet.getColumnWidth(location.col))
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
      // `*x` fill: expand to the column width (accounting's left-pinned $).
      if (
        typeof raw === 'number' &&
        typeof pattern === 'string' &&
        pattern.includes('*') &&
        style?.tb !== WrapStrategy.WRAP &&
        !location.worksheet.getMergedCell(location.row, location.col)
      ) {
        const fontString = getFontStyleString(style ?? undefined).fontString
        const measure = (text: string) => FontCache.getMeasureText(text, fontString).width
        const width = location.worksheet.getColumnWidth(location.col)
        const expanded = expandAsteriskFill(pattern, raw, width, measure)
        if (expanded !== null && expanded !== String(cell.v)) {
          // No hash pass: the fill was sized to the column, so the value
          // fits by construction (a too-wide value returns null above and
          // takes the normal path, hashes included).
          return next({ ...cell, v: expanded, t: CellValueType.NUMBER })
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
}
