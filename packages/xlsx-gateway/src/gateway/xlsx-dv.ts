/// Declarative data-validation save: the renderer snapshots the full Univer
/// rule set of a dirty sheet and this module rewrites the worksheet's
/// `<dataValidations>` section from it (mirroring the CF/filter recipe).
/// Mappings are the exact inverse of the read-side install in App.tsx.
/// Rules only xlsx's 2009 x14 extension can hold (a formula referencing
/// another sheet) ride in the worksheet extLst's x14:dataValidations block
/// instead, paired with nothing in the legacy section — the fallback shape
/// Excel itself writes for those rules.

import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '../shared/grid-bounds'

export class DvEditError extends Error {}

export interface DvCellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/// One rule in the Univer data-validation model shape (validated structurally
/// here — unknown shapes fail the save rather than guess).
export interface DvWireRule {
  readonly ranges: readonly DvCellArea[]
  readonly rule: Record<string, unknown>
}

const DV_TYPES = new Set(['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'])
const DV_OPERATORS = new Set([
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])
/// Univer DataValidationErrorStyle: INFO=0, STOP=1 (OOXML default), WARNING=2.
const DV_ERROR_STYLE_NAMES: Record<number, string | undefined> = {
  0: 'information',
  1: undefined,
  2: 'warning',
}

/// The worksheet extLst extension carrying x14 data validations. Excel
/// declares it with this uri; the {78C0D931-…} extension is the x14
/// conditional-formatting one, and Excel ignores validation rules parked
/// there — so the block gets an extension of its own.
const X14_DV_EXT_URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}'
const X14_XMLNS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM_XMLNS = 'http://schemas.microsoft.com/office/excel/2006/main'
const X14_DV_SECTION_RE =
  /<x14:dataValidations\b[^>]*>[\s\S]*?<\/x14:dataValidations>|<x14:dataValidations\b[^>]*\/>/g
const X14_DV_ENTRY_RE =
  /<x14:dataValidation\b[^>]*>[\s\S]*?<\/x14:dataValidation>|<x14:dataValidation\b[^>]*\/>/g
const X14_DV_EXT_OPEN_RE = new RegExp(
  `<ext\\b[^>]*\\buri="${X14_DV_EXT_URI.replace(/[{}]/g, '\\$&')}"[^>]*>`,
  'i',
)
const EXT_OPEN_RE = /<ext\b[^>]*>/g

export interface DvApplyOptions {
  /** keep the existing rules, replace those on the same ranges, drop `remove`, add the rest */
  readonly append?: boolean | undefined
  readonly remove?: readonly DvCellArea[] | undefined
}

export function applyDvRules(
  worksheetXml: string,
  rules: readonly DvWireRule[],
  options: DvApplyOptions = {},
): string {
  if (options.append) return appendDvRules(worksheetXml, rules, options.remove ?? [])
  const prepared = rules.map(prepareRule)
  const legacy = prepared.filter((rule) => !rule.needsX14)
  const x14 = prepared.filter((rule) => rule.needsX14)
  let xml = worksheetXml.replace(
    /<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>|<dataValidations\b[^>]*\/>/g,
    '',
  )
  if (legacy.length > 0) {
    xml = insertDvSection(
      xml,
      `<dataValidations count="${legacy.length}">${legacy.map(legacyEntry).join('')}</dataValidations>`,
    )
  }
  return reconcileX14DataValidations(xml, x14.map(x14Entry).join(''))
}

const DV_SECTION_RE =
  /<dataValidations\b[^>]*>([\s\S]*?)<\/dataValidations>|<dataValidations\b[^>]*\/>/
const DV_ENTRY_RE = /<dataValidation\b[^>]*?\/>|<dataValidation\b[^>]*>[\s\S]*?<\/dataValidation>/g

function appendDvRules(
  worksheetXml: string,
  rules: readonly DvWireRule[],
  remove: readonly DvCellArea[],
): string {
  const prepared = rules.map(prepareRule)
  const section = DV_SECTION_RE.exec(worksheetXml)
  const replaced = new Set(
    [...rules.flatMap((rule) => rule.ranges), ...remove].map((area) => normalizeRef(toRef(area))),
  )
  // an existing rule loses the areas the batch takes over; a multi-area sqref keeps the rest
  const kept: string[] = []
  for (const entry of section?.[1] ? [...section[1].matchAll(DV_ENTRY_RE)].map((m) => m[0]) : []) {
    const sqref = /\bsqref="([^"]*)"/.exec(entry)?.[1] ?? ''
    const areas = sqref.split(/\s+/).filter(Boolean)
    const remaining = areas.filter((area) => !replaced.has(normalizeRef(area)))
    if (remaining.length === 0) continue
    kept.push(
      remaining.length === areas.length
        ? entry
        : // function replacer: a surviving area's text is document-controlled and a
          // a string replacement would expand $& / $1 in the surviving area
          entry.replace(/\bsqref="[^"]*"/, () => `sqref="${remaining.join(' ')}"`),
    )
  }
  const entries = [...kept, ...prepared.filter((rule) => !rule.needsX14).map(legacyEntry)]
  let xml = section ? worksheetXml.replace(section[0], '') : worksheetXml
  if (entries.length > 0) {
    const body = `<dataValidations count="${entries.length}">${entries.join('')}</dataValidations>`
    xml = section
      ? xml.slice(0, section.index) + body + xml.slice(section.index)
      : insertDvSection(xml, body)
  }
  // Same takeover semantics for the file's x14 entries: they lose the areas
  // the batch claims and keep the rest verbatim; x14-only additions append.
  const x14Entries = [
    ...survivingX14Entries(worksheetXml, replaced),
    ...prepared.filter((rule) => rule.needsX14).map(x14Entry),
  ]
  return reconcileX14DataValidations(xml, x14Entries.join(''))
}

/** `A1:A1` and `$A$1` name the same cell as `A1`; `A:A` / `1:1` expand to the full grid. */
function normalizeRef(ref: string): string {
  const [a, b] = ref.replace(/\$/g, '').split(':')
  if (b !== undefined) {
    if (/^[A-Z]+$/.test(a!) && /^[A-Z]+$/.test(b)) return `${a}1:${b}${MAX_GRID_ROWS}`
    if (/^\d+$/.test(a!) && /^\d+$/.test(b))
      return `A${a}:${columnToLetters(MAX_GRID_COLUMNS - 1)}${b}`
  }
  return b === undefined || b === a ? a! : `${a}:${b}`
}

function insertDvSection(xml: string, section: string): string {
  const anchor =
    /<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/.exec(
      xml,
    )
  if (anchor) return xml.slice(0, anchor.index) + section + xml.slice(anchor.index)
  const end = xml.lastIndexOf('</worksheet>')
  if (end === -1) throw new DvEditError('Worksheet has no closing element.')
  return xml.slice(0, end) + section + xml.slice(end)
}

/// One wire rule after validation and normalization, holding the OOXML-level
/// formula texts once (shared by the legacy and x14 serializers) and the
/// verdict on which of the two shapes can carry it.
interface PreparedRule {
  readonly ranges: readonly DvCellArea[]
  /** OOXML type; undefined is the attribute-less "any" form */
  readonly type: string | undefined
  readonly rule: Record<string, unknown>
  readonly formula1: string | undefined
  readonly formula2: string | undefined
  readonly needsX14: boolean
}

function prepareRule(wireRule: DvWireRule): PreparedRule {
  if (wireRule.ranges.length === 0) {
    throw new DvEditError('A data-validation rule has no ranges.')
  }
  let rule = wireRule.rule
  let rawType = String(rule.type ?? '')
  if (rawType === 'listMultiple') {
    throw new DvEditError(
      'Multi-select list rules are Univer-only and cannot be saved to xlsx — ' +
        'delete the rule before saving.',
    )
  }
  if (rawType === 'checkbox') {
    // Checkbox is Univer-only; degrade to a two-value list so Excel keeps the
    // constraint. The default 1/0 pair round-trips back to a checkbox on load.
    const checked =
      rule.formula1 === undefined || rule.formula1 === '' ? '1' : String(rule.formula1)
    const unchecked =
      rule.formula2 === undefined || rule.formula2 === '' ? '0' : String(rule.formula2)
    rule = {
      ...rule,
      type: 'list',
      operator: undefined,
      formula1: `${checked},${unchecked}`,
      formula2: undefined,
    }
    rawType = 'list'
  }
  // 'any' is the read-side mapping of OOXML type="none" (no constraint, just
  // messages); it round-trips back to the default, attribute-less form.
  const type = rawType === 'any' || rawType === 'none' ? undefined : rawType
  if (type !== undefined && !DV_TYPES.has(type)) {
    throw new DvEditError(`Unsupported data-validation type "${rawType}".`)
  }
  const operator = rule.operator === undefined ? undefined : String(rule.operator)
  if (operator !== undefined && operator !== '') {
    if (!DV_OPERATORS.has(operator)) {
      throw new DvEditError(`Unsupported data-validation operator "${operator}".`)
    }
  }
  errorStyleName(rule.errorStyle)
  const formula1 = formulaText(type, rule.formula1)
  const formula2 = formulaText(type, rule.formula2)
  return {
    ranges: wireRule.ranges,
    type,
    rule,
    formula1,
    formula2,
    needsX14: needsX14Extension(type, rule, formula1, formula2),
  }
}

/// The one feature the 2006 legacy grammar cannot hold: a formula qualified
/// with a sheet name. Excel parks exactly those rules in the x14 extension
/// (leaving no legacy twin), so the save matches that shape. A list literal
/// is only ever a reference when the model value starts with `=`.
function needsX14Extension(
  type: string | undefined,
  rule: Record<string, unknown>,
  formula1: string | undefined,
  formula2: string | undefined,
): boolean {
  if (type === 'list') {
    return String(rule.formula1 ?? '').startsWith('=') && sheetQualified(formula1)
  }
  return sheetQualified(formula1) || sheetQualified(formula2)
}

/// A `!` behind a sheet name (quoted, bare, or a 3-D pair) qualifies the
/// reference onto another sheet; an in-sheet formula has no such token.
const SHEET_QUALIFIED_RE =
  /(?:^|[^A-Za-z0-9_.])(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)(?::(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*))?!/

function sheetQualified(formula: string | undefined): boolean {
  return formula !== undefined && SHEET_QUALIFIED_RE.test(formula)
}

/// Attributes shared verbatim by the legacy and x14 entry forms.
function entryAttributes(prepared: PreparedRule): string[] {
  const attrs: string[] = []
  const { type, rule } = prepared
  if (type !== undefined) attrs.push(`type="${type}"`)
  const operator = rule.operator === undefined ? undefined : String(rule.operator)
  if (operator !== undefined && operator !== '') {
    // "between" is the OOXML default; only operator-carrying types keep it.
    if (operator !== 'between' && type !== undefined && type !== 'list' && type !== 'custom') {
      attrs.push(`operator="${operator}"`)
    }
  }
  if (rule.allowBlank === true) attrs.push('allowBlank="1"')
  // OOXML's showDropDown="1" SUPPRESSES the in-cell dropdown (inverted name);
  // Univer's showDropDown means what it says.
  if (type === 'list' && rule.showDropDown === false) attrs.push('showDropDown="1"')
  if (rule.showInputMessage === true) attrs.push('showInputMessage="1"')
  if (rule.showErrorMessage === true) attrs.push('showErrorMessage="1"')
  const errorStyle = errorStyleName(rule.errorStyle)
  if (errorStyle !== undefined) attrs.push(`errorStyle="${errorStyle}"`)
  for (const [key, attribute] of [
    ['errorTitle', 'errorTitle'],
    ['error', 'error'],
    ['promptTitle', 'promptTitle'],
    ['prompt', 'prompt'],
  ] as const) {
    const value = rule[key]
    if (typeof value === 'string' && value.length > 0) {
      attrs.push(`${attribute}="${escapeXmlAttribute(value)}"`)
    }
  }
  return attrs
}

function legacyEntry(prepared: PreparedRule): string {
  const attrs = [...entryAttributes(prepared), `sqref="${prepared.ranges.map(toRef).join(' ')}"`]
  const formulas = legacyFormulas(prepared)
  return formulas === ''
    ? `<dataValidation ${attrs.join(' ')}/>`
    : `<dataValidation ${attrs.join(' ')}>${formulas}</dataValidation>`
}

function legacyFormulas(prepared: PreparedRule): string {
  return (
    (prepared.formula1 === undefined
      ? ''
      : `<formula1>${escapeXmlText(prepared.formula1)}</formula1>`) +
    (prepared.formula2 === undefined
      ? ''
      : `<formula2>${escapeXmlText(prepared.formula2)}</formula2>`)
  )
}

/// The x14 twin: same attributes, formulas wrapped one level deeper in
/// <xm:f>, and the ranges as an <xm:sqref> child instead of an attribute.
function x14Entry(prepared: PreparedRule): string {
  const formulas =
    (prepared.formula1 === undefined
      ? ''
      : `<x14:formula1><xm:f>${escapeXmlText(prepared.formula1)}</xm:f></x14:formula1>`) +
    (prepared.formula2 === undefined
      ? ''
      : `<x14:formula2><xm:f>${escapeXmlText(prepared.formula2)}</xm:f></x14:formula2>`)
  const attrs = entryAttributes(prepared).join(' ')
  return (
    `<x14:dataValidation ${attrs}>${formulas}` +
    `<xm:sqref>${prepared.ranges.map(toRef).join(' ')}</xm:sqref></x14:dataValidation>`
  )
}

/// Rewrites the worksheet's x14:dataValidations block to `entriesXml` (empty
/// removes it). The block lives in its own {CCE6A557} extension — Excel's
/// x14 data-validation extension — never in the {78C0D931} one, which
/// carries conditional formatting and whose validation rules Excel ignores.
/// An extension left without elements and an extLst left without extensions
/// go away so deleted rules leave no husks.
function reconcileX14DataValidations(xml: string, entriesXml: string): string {
  if (entriesXml === '' && !/<x14:dataValidations\b/.test(xml)) return xml
  xml = removeX14DvBlocks(xml)
  if (entriesXml === '') return dropEmptyX14DvExt(xml)
  const count = (entriesXml.match(/<x14:dataValidation\b/g) ?? []).length
  const block =
    `<x14:dataValidations xmlns:x14="${X14_XMLNS}" xmlns:xm="${XM_XMLNS}" count="${count}">` +
    `${entriesXml}</x14:dataValidations>`
  const extOpen = X14_DV_EXT_OPEN_RE.exec(xml)
  if (extOpen) {
    if (extOpen[0].endsWith('/>')) {
      return (
        xml.slice(0, extOpen.index) +
        `<ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_XMLNS}">${block}</ext>` +
        xml.slice(extOpen.index + extOpen[0].length)
      )
    }
    // The extension holds the block and nothing else Excel writes there, so
    // the block goes in right before the extension closes.
    const close = xml.indexOf('</ext>', extOpen.index)
    if (close === -1) throw new DvEditError('The data-validation extension element is malformed.')
    return xml.slice(0, close) + block + xml.slice(close)
  }
  // No {CCE6A557} extension yet: only the worksheet-level extLst (past
  // sheetData) qualifies — cells may carry their own extLst.
  const ext = `<ext uri="${X14_DV_EXT_URI}" xmlns:x14="${X14_XMLNS}">${block}</ext>`
  const extLstClose = xml.lastIndexOf('</extLst>')
  if (extLstClose !== -1 && extLstClose > xml.lastIndexOf('</sheetData>')) {
    return xml.slice(0, extLstClose) + ext + xml.slice(extLstClose)
  }
  const worksheetClose = xml.lastIndexOf('</worksheet>')
  if (worksheetClose === -1) throw new DvEditError('Worksheet has no closing element.')
  return xml.slice(0, worksheetClose) + `<extLst>${ext}</extLst>` + xml.slice(worksheetClose)
}

/// One worksheet extLst extension element: its open tag, its content and
/// the `</ext>` closing it. ext elements never nest.
interface X14Ext {
  readonly start: number
  readonly openEnd: number
  readonly end: number
}

/// The ext element holding `index`, if any: the last `<ext …>` open tag
/// before it whose `</ext>` still comes after.
function extAround(xml: string, index: number): X14Ext | null {
  let found: X14Ext | null = null
  for (const match of xml.matchAll(EXT_OPEN_RE)) {
    const start = match.index ?? 0
    if (start >= index) break
    if (match[0].endsWith('/>')) continue
    const openEnd = start + match[0].length
    const close = xml.indexOf('</ext>', openEnd)
    if (close === -1) continue
    const end = close + '</ext>'.length
    if (end <= index) continue
    found = { start, openEnd, end }
  }
  return found
}

/// Removes the worksheet's x14:dataValidations blocks. Excel keeps the block
/// in the {CCE6A557} extension, but a file written by an older build may
/// hold it in the conditional-formatting one, whose rules Excel ignores; the
/// extension the removal empties goes with it, since an extension without
/// elements is what Excel flags for repair.
function removeX14DvBlocks(xml: string): string {
  let out = xml
  for (const block of [...out.matchAll(X14_DV_SECTION_RE)].reverse()) {
    const index = block.index ?? 0
    const ext = extAround(out, index)
    const removed = block[0].length
    out = out.slice(0, index) + out.slice(index + removed)
    if (ext === null) continue
    // The block sat inside the extension, so its end moved up by its size.
    const end = ext.end - removed
    if (out.slice(ext.openEnd, end - '</ext>'.length).includes('<')) continue
    out = dropExt(out, { start: ext.start, openEnd: ext.openEnd, end })
  }
  return out
}

/// Removes the ext element and the extLst around it when that was its only
/// extension — an extLst without extensions is what Excel flags for repair.
function dropExt(xml: string, ext: X14Ext): string {
  const stripped = xml.slice(0, ext.start) + xml.slice(ext.end)
  // The enclosing extLst (a malformed worksheet may leave the ext bare);
  // drop it too when this was its only extension.
  const lstOpenAt = xml.lastIndexOf('<extLst', ext.start)
  if (lstOpenAt === -1) return stripped
  const lstOpen = /<extLst\b[^>]*>/.exec(xml.slice(lstOpenAt, ext.start))?.[0]
  if (lstOpen === undefined) return stripped
  const lstInnerStart = lstOpenAt + lstOpen.length
  const lstClose = xml.indexOf('</extLst>', lstInnerStart)
  if (lstClose === -1 || lstClose < ext.start) return stripped
  const inner = xml.slice(lstInnerStart, ext.start) + xml.slice(ext.end, lstClose)
  if (inner.includes('<')) return stripped
  return xml.slice(0, lstOpenAt) + xml.slice(lstClose + '</extLst>'.length)
}

/// Removes the {CCE6A557} extension when it holds no elements, and the
/// extLst when that in turn left it without extensions.
function dropEmptyX14DvExt(xml: string): string {
  const extOpen = X14_DV_EXT_OPEN_RE.exec(xml)
  if (!extOpen || extOpen[0].endsWith('/>')) return xml
  const close = xml.indexOf('</ext>', extOpen.index)
  if (close === -1) return xml
  const start = extOpen.index
  if (xml.slice(start + extOpen[0].length, close).includes('<')) return xml
  return dropExt(xml, {
    start,
    openEnd: start + extOpen[0].length,
    end: close + '</ext>'.length,
  })
}

/// The file's own x14 entries for append mode: each loses the areas the batch
/// takes over (dropped when none remain, re-ranged otherwise) and keeps the
/// rest of its bytes verbatim — the same discipline as the legacy section.
function survivingX14Entries(xml: string, replaced: ReadonlySet<string>): string[] {
  const kept: string[] = []
  for (const match of xml.matchAll(X14_DV_ENTRY_RE)) {
    const entry = match[0]
    const sqref = /<xm:sqref>([^<]*)<\/xm:sqref>/.exec(entry)?.[1] ?? ''
    const areas = sqref.split(/\s+/).filter(Boolean)
    const remaining = areas.filter((area) => !replaced.has(normalizeRef(area)))
    if (remaining.length === 0) continue
    kept.push(
      remaining.length === areas.length
        ? entry
        : // function replacer: the surviving areas are document-controlled and
          // a string replacement would expand $& / $1 in them
          entry.replace(
            /<xm:sqref>[^<]*<\/xm:sqref>/,
            () => `<xm:sqref>${remaining.join(' ')}</xm:sqref>`,
          ),
    )
  }
  return kept
}

function errorStyleName(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined
  const style = Number(value)
  if (!(style in DV_ERROR_STYLE_NAMES)) {
    throw new DvEditError(`Unsupported data-validation error style "${String(value)}".`)
  }
  return DV_ERROR_STYLE_NAMES[style]
}

/// Inverse of the install-side formula transforms: list literals regain their
/// quotes, `=`-prefixed references/formulas lose the prefix, and panel-edited
/// date/time strings become Excel serial numbers.
function formulaText(type: string | undefined, raw: unknown): string | undefined {
  if (raw === undefined || raw === null) return undefined
  const text = String(raw)
  if (text === '') return undefined
  if (type === 'list') {
    return text.startsWith('=') ? text.slice(1) : `"${text.replaceAll('"', '""')}"`
  }
  if (type === 'custom') {
    return text.startsWith('=') ? text.slice(1) : text
  }
  if (type === 'date') {
    const serial = dateToSerial(text)
    if (serial !== undefined) return String(serial)
  }
  if (type === 'time') {
    const fraction = timeToFraction(text)
    if (fraction !== undefined) return String(fraction)
  }
  return text.startsWith('=') ? text.slice(1) : text
}

/// 'YYYY-MM-DD[ HH:mm[:ss]]' (or slashes) → Excel serial (days since
/// 1899-12-30). Plain numbers and references pass through untouched.
/// Impossible calendar dates or clock times return undefined so the caller
/// keeps the original text instead of writing a silently wrong serial; a
/// pre-1900 year throws DvEditError, because it has no serial to write.
/// Serials 1-59 sit one day before the linear rule (Excel's phantom 1900-02-29);
/// formatSerial in @genoffice/file-parse reverses the same shift.
function dateToSerial(text: string): number | undefined {
  const match =
    /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/.exec(
      text.trim(),
    )
  if (!match) return undefined
  const [, year, month, day, hour, minute, second] = match
  const yearNum = Number(year)
  const monthNum = Number(month)
  const dayNum = Number(day)
  // Excel's 1900 date system has no serial before 1900-01-01, so such a date
  // cannot be written at all. Date.UTC also maps a 0..99 year onto 19xx
  // (0099 -> 1999), which would otherwise store a plausible, silently wrong
  // serial; reject the year outright instead.
  if (yearNum < 1900) {
    throw new DvEditError(
      `Data-validation date "${text.trim()}" is before 1900, which Excel cannot store.`,
    )
  }
  if (monthNum < 1 || monthNum > 12) return undefined
  if (dayNum < 1 || dayNum > daysInMonth(yearNum, monthNum)) return undefined
  let seconds = 0
  if (hour !== undefined) {
    const hourNum = Number(hour)
    const minuteNum = Number(minute)
    const secondNum = Number(second ?? 0)
    if (hourNum < 0 || hourNum > 23) return undefined
    if (minuteNum < 0 || minuteNum > 59) return undefined
    if (secondNum < 0 || secondNum > 59) return undefined
    seconds = hourNum * 3600 + minuteNum * 60 + secondNum
  }
  const days = (Date.UTC(yearNum, monthNum - 1, dayNum) - Date.UTC(1899, 11, 30)) / 86_400_000
  const serial = days < 61 ? days - 1 : days
  return seconds === 0 ? serial : serial + seconds / 86_400
}

/// Days in a 1-based month, with the Gregorian leap-year rule for February.
function daysInMonth(year: number, month: number): number {
  if (month === 2) return isLeapYear(year) ? 29 : 28
  if (month === 4 || month === 6 || month === 9 || month === 11) return 30
  return 31
}

/// Gregorian leap-year rule: divisible by 4, except centuries not by 400.
function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0
}

function timeToFraction(text: string): number | undefined {
  const match = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(text.trim())
  if (!match) return undefined
  const [, hour, minute, second] = match
  const hourNum = Number(hour)
  const minuteNum = Number(minute)
  const secondNum = Number(second ?? 0)
  if (hourNum < 0 || hourNum > 23) return undefined
  if (minuteNum < 0 || minuteNum > 59) return undefined
  if (secondNum < 0 || secondNum > 59) return undefined
  return (hourNum * 3600 + minuteNum * 60 + secondNum) / 86_400
}

function toRef(range: DvCellArea): string {
  if (
    range.startRow < 0 ||
    range.startColumn < 0 ||
    range.endRow < range.startRow ||
    range.endColumn < range.startColumn ||
    range.endRow >= MAX_GRID_ROWS ||
    range.endColumn >= MAX_GRID_COLUMNS
  ) {
    throw new DvEditError('A data-validation range lies outside the worksheet grid.')
  }
  const start = `${columnToLetters(range.startColumn)}${range.startRow + 1}`
  return range.startRow === range.endRow && range.startColumn === range.endColumn
    ? start
    : `${start}:${columnToLetters(range.endColumn)}${range.endRow + 1}`
}

function columnToLetters(column: number): string {
  let letters = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return letters
}

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll('\n', '&#10;')
}
