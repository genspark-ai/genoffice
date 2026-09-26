import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { splitSheetRef } from '@genoffice/xlsx-gateway/domain/chart-visual'
import { parseRelationships, parseSheetElements } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import {
  FORMULA_REFERENCE_PATTERN,
  qualifierMatches,
} from '@genoffice/xlsx-gateway/gateway/xlsx-structure'
import JSZip from 'jszip'
import { excerpt, PLACEHOLDER, textWidthPx, type IssueDraft } from '../check'
import { displayText } from './xlsx'

export const SHEET_CHECKS = [
  'formula_error',
  'formula_not_evaluated',
  'missing_sheet_ref',
  'defined_name_broken',
  'chart_ref',
  'number_overflow',
  'placeholder_left',
  'rules_present',
] as const

const PX_PER_CHAR = 7
/** column padding Excel adds around the digits */
const CELL_PADDING_PX = 5
const LIST_CAP = 20

interface Cell {
  ref: string
  column: number
  row: number
  styleIndex: number
  type: string | undefined
  formula: string | undefined
  hasFormula: boolean
  value: string | undefined
  inline: string | undefined
}

interface Xf {
  numFmtId: number
  fontId: number
  wrap: boolean
}

interface Styles {
  numFmts: Map<number, string>
  fontSizes: number[]
  xfs: Xf[]
}

interface SheetXml {
  name: string
  path: string
  xml: string
}

export interface WorkbookCheck {
  issues: IssueDraft[]
  sheets: string[]
}

export async function checkWorkbook(bytes: Uint8Array): Promise<WorkbookCheck> {
  const zip = await JSZip.loadAsync(bytes)
  const text = async (path: string) => (await zip.file(path)?.async('string')) ?? ''
  const workbook = await text('xl/workbook.xml')
  const rels = parseRelationships(await text('xl/_rels/workbook.xml.rels'))
  const targets = new Map(
    rels.filter((r) => r.id !== undefined && !r.external).map((r) => [r.id!, partPath(r.target)]),
  )
  const sheets: SheetXml[] = []
  for (const el of parseSheetElements(workbook)) {
    const path = el.relationshipId ? targets.get(el.relationshipId) : undefined
    if (path) sheets.push({ name: el.name, path, xml: await text(path) })
  }
  const names = sheets.map((s) => s.name)
  const styles = parseStyles(await text('xl/styles.xml'))
  const shared = parseSharedStrings(await text('xl/sharedStrings.xml'))
  const date1904 = /<workbookPr\b[^>]*\bdate1904="(?:1|true)"/.test(workbook)
  const issues: IssueDraft[] = []

  for (const sheet of sheets) {
    const cells = parseCells(sheet.xml)
    checkFormulas(sheet.name, cells, names, issues)
    checkOverflow(sheet, cells, styles, shared, date1904, issues)
    checkPlaceholders(sheet.name, cells, shared, issues)
    checkRulesFormulas(sheet, names, issues)
  }
  checkDefinedNames(workbook, names, issues)
  for (const path of Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d*\.xml$/.test(f))) {
    checkChart(path, await text(path), sheets, issues)
  }
  return { issues, sheets: names }
}

// ── parsing ────────────────────────────────────────────────────────────

function partPath(target: string): string {
  return target.startsWith('/') ? target.slice(1) : `xl/${target}`
}

function unescapeXml(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(?:x([0-9a-fA-F]+)|(\d+));/g, (reference, hex: string | undefined, dec: string) => {
      const code = hex === undefined ? Number(dec) : parseInt(hex, 16)
      return code <= 0x10ffff && (code < 0xd800 || code > 0xdfff)
        ? String.fromCodePoint(code)
        : reference
    })
    .replace(/&amp;/g, '&')
}

function parseStyles(xml: string): Styles {
  const numFmts = new Map<number, string>()
  for (const m of xml.matchAll(/<numFmt\b([^>]*)\/?>/g)) {
    const id = /\bnumFmtId="(\d+)"/.exec(m[1]!)?.[1]
    const code = /\bformatCode="([^"]*)"/.exec(m[1]!)?.[1]
    if (id !== undefined && code !== undefined) numFmts.set(Number(id), unescapeXml(code))
  }
  const fontSizes: number[] = []
  const fonts = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml)?.[1] ?? ''
  for (const m of fonts.matchAll(/<font\b[^>]*(?:\/>|>([\s\S]*?)<\/font>)/g)) {
    fontSizes.push(Number(/<sz\b[^>]*\bval="([\d.]+)"/.exec(m[1] ?? '')?.[1] ?? 11))
  }
  const xfs: Xf[] = []
  const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1] ?? ''
  for (const m of cellXfs.matchAll(/<xf\b([^>]*)(?:\/>|>([\s\S]*?)<\/xf>)/g)) {
    xfs.push({
      numFmtId: Number(/\bnumFmtId="(\d+)"/.exec(m[1]!)?.[1] ?? 0),
      fontId: Number(/\bfontId="(\d+)"/.exec(m[1]!)?.[1] ?? 0),
      wrap: /<alignment\b[^>]*\bwrapText="(?:1|true)"/.test(m[2] ?? ''),
    })
  }
  return { numFmts, fontSizes, xfs }
}

function parseSharedStrings(xml: string): string[] {
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => plainText(m[1]!))
}

function plainText(richXml: string): string {
  return [...richXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => unescapeXml(t[1]!)).join('')
}

function parseCells(xml: string): Cell[] {
  const cells: Cell[] = []
  const data = /<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml)?.[1] ?? ''
  for (const m of data.matchAll(/<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g)) {
    const attrs = m[1]!
    const ref = /\br="([A-Z]+\d+)"/.exec(attrs)?.[1]
    if (!ref) continue
    const body = m[3] ?? ''
    const f = /<f\b([^>]*)(?:\/>|>([\s\S]*?)<\/f>)/.exec(body)
    const v = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1]
    const is = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(body)?.[1]
    const bounds = parseRange(ref)
    cells.push({
      ref,
      column: bounds.startColumn,
      row: bounds.startRow,
      styleIndex: Number(/\bs="(\d+)"/.exec(attrs)?.[1] ?? 0),
      type: /\bt="([^"]+)"/.exec(attrs)?.[1],
      formula: f?.[2] !== undefined ? unescapeXml(f[2]) : undefined,
      hasFormula: f !== null,
      value: v === undefined ? undefined : unescapeXml(v),
      inline: is === undefined ? undefined : plainText(is),
    })
  }
  return cells
}

// ── formulas ───────────────────────────────────────────────────────────

/** Sheet qualifiers of every A1 reference in a formula, string literals skipped. */
function sheetQualifiers(formula: string): string[] {
  const out: string[] = []
  formula.split('"').forEach((segment, index) => {
    if (index % 2 === 1) return
    for (const m of segment.matchAll(FORMULA_REFERENCE_PATTERN)) {
      if (m[2] !== undefined) out.push(m[2])
    }
  })
  return out
}

/**
 * `[1]Sheet1!A1`, `'[Book.xlsx]Sheet 1'!A1` and `'C:\\path\\[Book.xlsx]Sheet'!A1` point at another
 * workbook, not a local sheet; Excel forbids `[` in sheet names, so any qualifier holding one is external.
 * The quoted form matches one quote-delimited token (no `'` or `!` inside), so it cannot start at the
 * closing quote of an earlier local qualifier and run on to a later external one.
 */
const EXTERNAL_REF = /'[^'!]*\[[^'!]*'!|\[\d+\][^!\s(),]*!/g

function missingSheets(formula: string, names: readonly string[]): string[] {
  const missing = new Set<string>()
  for (const q of sheetQualifiers(formula.replace(EXTERNAL_REF, ''))) {
    if (!names.some((n) => sameSheet(q, n))) missing.add(q)
  }
  return [...missing]
}

/** Excel compares sheet names case-insensitively. */
function sameSheet(qualifier: string, name: string): boolean {
  return qualifierMatches(qualifier.toLowerCase(), name.toLowerCase())
}

function checkFormulas(
  sheet: string,
  cells: Cell[],
  names: readonly string[],
  issues: IssueDraft[],
): void {
  const unevaluated: string[] = []
  for (const c of cells) {
    if (!c.hasFormula) continue
    const at = `${sheet}!${c.ref}`
    if (c.type === 'e' && c.value) {
      issues.push({
        code: 'formula_error',
        level: 'error',
        path: at,
        sheet,
        cell: c.ref,
        message: `formula evaluates to ${c.value}`,
        ...(c.formula ? { context: `=${excerpt(c.formula)}` } : {}),
      })
    } else if (c.value === undefined) {
      unevaluated.push(c.ref)
    }
    if (c.formula) {
      for (const q of missingSheets(c.formula, names)) {
        issues.push({
          code: 'missing_sheet_ref',
          level: 'error',
          path: at,
          sheet,
          cell: c.ref,
          message: `formula references a sheet that does not exist: ${q}`,
          context: `=${excerpt(c.formula)}`,
        })
      }
    }
  }
  if (unevaluated.length) {
    issues.push({
      code: 'formula_not_evaluated',
      level: 'warning',
      path: `${sheet}!${unevaluated[0]}`,
      sheet,
      cell: unevaluated[0]!,
      message: `${unevaluated.length} formula(s) have no cached value; Excel computes them on open, readers without a calc engine see empty cells`,
      context: listCells(unevaluated),
    })
  }
}

function listCells(refs: readonly string[]): string {
  return refs.length > LIST_CAP
    ? `${refs.slice(0, LIST_CAP).join(', ')} … (${refs.length} cells)`
    : refs.join(', ')
}

/** Conditional-format and data-validation formulas live outside sheetData. */
function checkRulesFormulas(sheet: SheetXml, names: readonly string[], issues: IssueDraft[]): void {
  const cf = (sheet.xml.match(/<conditionalFormatting\b/g) ?? []).length
  const dv = (sheet.xml.match(/<dataValidation\b/g) ?? []).length
  if (cf || dv) {
    issues.push({
      code: 'rules_present',
      level: 'info',
      path: sheet.name,
      sheet: sheet.name,
      message: `${cf} conditional format(s), ${dv} data validation(s) already on the sheet; add_conditional_format and set_data_validation add to them (clear_conditional_formats starts over)`,
    })
  }
  for (const m of sheet.xml.matchAll(
    /<(?:cfRule|dataValidation)\b[^>]*\bsqref="([^"]*)"[^>]*>([\s\S]*?)<\/(?:cfRule|dataValidation)>|<conditionalFormatting\b[^>]*\bsqref="([^"]*)"[^>]*>([\s\S]*?)<\/conditionalFormatting>/g,
  )) {
    const range = m[1] ?? m[3] ?? ''
    const body = m[2] ?? m[4] ?? ''
    for (const f of body.matchAll(
      /<(?:formula|formula1|formula2)\b[^>]*>([\s\S]*?)<\/(?:formula|formula1|formula2)>/g,
    )) {
      const formula = unescapeXml(f[1]!)
      for (const q of missingSheets(formula, names)) {
        issues.push({
          code: 'missing_sheet_ref',
          level: 'error',
          path: `${sheet.name}!${range}`,
          sheet: sheet.name,
          range,
          message: `rule formula references a sheet that does not exist: ${q}`,
          context: `=${excerpt(formula)}`,
        })
      }
    }
  }
}

function checkDefinedNames(workbook: string, names: readonly string[], issues: IssueDraft[]): void {
  for (const m of workbook.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const name = unescapeXml(/\bname="([^"]*)"/.exec(m[1]!)?.[1] ?? '')
    const formula = unescapeXml(m[2]!)
    if (!name) continue
    const path = `definedName ${name}`
    if (/#REF!/.test(formula)) {
      issues.push({
        code: 'defined_name_broken',
        level: 'error',
        path,
        message: `defined name ${name} points at deleted cells`,
        context: `=${excerpt(formula)}`,
        suggest: { op: 'delete_defined_name', name },
      })
      continue
    }
    for (const q of missingSheets(formula, names)) {
      issues.push({
        code: 'missing_sheet_ref',
        level: 'error',
        path,
        message: `defined name ${name} references a sheet that does not exist: ${q}`,
        context: `=${excerpt(formula)}`,
      })
    }
  }
}

// ── charts ─────────────────────────────────────────────────────────────

function checkChart(path: string, xml: string, sheets: SheetXml[], issues: IssueDraft[]): void {
  const part = path.replace(/^xl\//, '')
  const seen = new Set<string>()
  for (const m of xml.matchAll(/<c:f>([\s\S]*?)<\/c:f>/g)) {
    const ref = unescapeXml(m[1]!).trim()
    if (seen.has(ref)) continue
    seen.add(ref)
    const split = splitSheetRef(ref)
    if (!split || split.sheetName.includes('[')) continue
    const sheet = sheets.find((s) => s.name.toLowerCase() === split.sheetName.toLowerCase())
    if (!sheet) {
      issues.push({
        code: 'chart_ref',
        level: 'error',
        path: part,
        part: path,
        message: `chart series references a sheet that does not exist: ${split.sheetName}`,
        context: ref,
      })
      continue
    }
    const dim = /<dimension\b[^>]*\bref="([^"]+)"/.exec(sheet.xml)?.[1]
    if (!dim) continue
    try {
      const used = parseRange(dim)
      const area = parseRange(split.range)
      const disjoint =
        area.startRow > used.endRow ||
        area.endRow < used.startRow ||
        area.startColumn > used.endColumn ||
        area.endColumn < used.startColumn
      const beyond =
        area.endRow > used.endRow ||
        area.endColumn > used.endColumn ||
        area.startRow < used.startRow ||
        area.startColumn < used.startColumn
      if (!disjoint && !beyond) continue
      issues.push({
        code: 'chart_ref',
        level: 'warning',
        path: part,
        part: path,
        message: disjoint
          ? `chart series range lies outside the used range of ${split.sheetName} (${dim}); the series plots empty`
          : `chart series range extends beyond the used range of ${split.sheetName} (${dim}); the cells outside it plot empty`,
        context: ref,
      })
    } catch {
      // whole-row/column refs are not chart series ranges
    }
  }
}

// ── grid text ──────────────────────────────────────────────────────────

interface ColumnWidths {
  defaultPx: number
  byColumn: Map<number, { px: number; hidden: boolean }>
}

function parseColumnWidths(xml: string): ColumnWidths {
  const fmt = /<sheetFormatPr\b([^>]*)\/?>/.exec(xml)?.[1] ?? ''
  const defaultChars =
    Number(/\bdefaultColWidth="([\d.]+)"/.exec(fmt)?.[1]) ||
    (Number(/\bbaseColWidth="(\d+)"/.exec(fmt)?.[1]) || 8) + 0.43
  const byColumn = new Map<number, { px: number; hidden: boolean }>()
  for (const m of xml.matchAll(/<col\b([^>]*)\/?>/g)) {
    const a = m[1]!
    const min = Number(/\bmin="(\d+)"/.exec(a)?.[1])
    const max = Number(/\bmax="(\d+)"/.exec(a)?.[1])
    if (!min || !max) continue
    const width = Number(/\bwidth="([\d.]+)"/.exec(a)?.[1])
    const hidden = /\bhidden="(?:1|true)"/.test(a)
    for (let c = min; c <= Math.min(max, min + 2000); c++) {
      byColumn.set(c - 1, {
        px: charsToPx(Number.isFinite(width) && width > 0 ? width : defaultChars),
        hidden,
      })
    }
  }
  return { defaultPx: charsToPx(defaultChars), byColumn }
}

function charsToPx(chars: number): number {
  return Math.round(chars * PX_PER_CHAR + CELL_PADDING_PX)
}

function columnLabel(index: number): string {
  let label = ''
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    label = String.fromCharCode(64 + ((n - 1) % 26) + 1) + label
  }
  return label
}

function checkOverflow(
  sheet: SheetXml,
  cells: Cell[],
  styles: Styles,
  shared: readonly string[],
  date1904: boolean,
  issues: IssueDraft[],
): void {
  const widths = parseColumnWidths(sheet.xml)
  const hiddenRows = new Set<number>()
  for (const m of sheet.xml.matchAll(/<row\b([^>]*)>/g)) {
    if (!/\bhidden="(?:1|true)"/.test(m[1]!)) continue
    const r = /\br="(\d+)"/.exec(m[1]!)?.[1]
    if (r !== undefined) hiddenRows.add(Number(r) - 1)
  }
  const merged = new Set<string>()
  for (const m of sheet.xml.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"/g)) {
    const r = parseRange(m[1]!)
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startColumn; col <= r.endColumn; col++) merged.add(`${row}:${col}`)
    }
  }
  const worst = new Map<number, { cell: string; text: string; needPx: number; count: number }>()
  for (const c of cells) {
    if (c.type && c.type !== 'n') continue
    if (c.value === undefined || c.value === '' || merged.has(`${c.row}:${c.column}`)) continue
    if (hiddenRows.has(c.row)) continue
    const number = Number(c.value)
    if (!Number.isFinite(number)) continue
    const xf = styles.xfs[c.styleIndex] ?? styles.xfs[0]
    if (xf?.wrap) continue
    const column = widths.byColumn.get(c.column)
    if (column?.hidden) continue
    const availablePx = (column?.px ?? widths.defaultPx) - CELL_PADDING_PX
    const numFmtId = xf?.numFmtId ?? 0
    const format = styles.numFmts.get(numFmtId) ?? builtinFormat(numFmtId)
    const sizePt = styles.fontSizes[xf?.fontId ?? 0] ?? 11
    // General shrinks decimals to fit; only the integer part is fixed
    const text =
      format === undefined || format === 'General'
        ? generalIntegerText(number)
        : displayText(number, format, date1904)
    const needPx = textWidthPx(text, sizePt)
    if (needPx <= availablePx + PX_PER_CHAR / 2) continue
    const entry = worst.get(c.column)
    if (!entry) worst.set(c.column, { cell: c.ref, text, needPx, count: 1 })
    else {
      entry.count++
      if (needPx > entry.needPx) Object.assign(entry, { cell: c.ref, text, needPx })
    }
  }
  for (const [col, entry] of worst) {
    const letter = columnLabel(col)
    const widthPx = (widths.byColumn.get(col)?.px ?? widths.defaultPx) - CELL_PADDING_PX
    const suggested = Math.ceil(entry.needPx) + PX_PER_CHAR
    issues.push({
      code: 'number_overflow',
      level: 'warning',
      path: `${sheet.name}!${letter}:${letter}`,
      sheet: sheet.name,
      range: `${letter}:${letter}`,
      cell: entry.cell,
      message: `${entry.count} numeric cell(s) in column ${letter} show ### (e.g. ${entry.cell} "${entry.text}" needs about ${Math.ceil(entry.needPx)} px, the column is ${widthPx} px)`,
      suggest: { op: 'set_col_width', sheet: sheet.name, column: letter, widthPx: suggested },
    })
  }
}

function generalIntegerText(value: number): string {
  const abs = Math.abs(value)
  if (abs >= 1e11) return value < 0 ? '-1.00000E+11' : '1.00000E+11'
  return String(Math.trunc(value))
}

const BUILTIN_FORMATS: Record<number, string> = {
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
}

function builtinFormat(id: number): string | undefined {
  return id === 0 ? 'General' : BUILTIN_FORMATS[id]
}

// ── text ───────────────────────────────────────────────────────────────

function checkPlaceholders(
  sheet: string,
  cells: Cell[],
  shared: readonly string[],
  issues: IssueDraft[],
): void {
  const hits: { ref: string; text: string }[] = []
  for (const c of cells) {
    const text =
      c.type === 's' && c.value !== undefined
        ? shared[Number(c.value)]
        : c.type === 'inlineStr'
          ? c.inline
          : c.type === 'str'
            ? c.value
            : undefined
    if (text && PLACEHOLDER.test(text)) hits.push({ ref: c.ref, text })
  }
  if (!hits.length) return
  issues.push({
    code: 'placeholder_left',
    level: 'warning',
    path: `${sheet}!${hits[0]!.ref}`,
    sheet,
    cell: hits[0]!.ref,
    message: `${hits.length} cell(s) contain placeholder text (template keys, TODO/TBD, lorem ipsum)`,
    context: listCells(hits.map((h) => `${h.ref} "${excerpt(h.text, 40)}"`)),
  })
}
