/// Lays the active sheet out as print HTML from the live Univer model —
/// display strings (number formats applied), cell styles, merges, and the
/// sheet's Page Layout settings (print area, repeated title rows, gridlines,
/// headings). The main process turns the HTML into a PDF.

import { htmlLang, type Lang } from '@genoffice/i18n'
import { columnIndex, columnLabel } from '../domain/cell-address'

import type { WorkbookExportPdfRequest } from '../shared/desktop-api'
import type { HeaderFooterParts, PageSetupJournalState } from './edit-journal'
import { getLang, t } from './i18n/locale'

export class PrintError extends Error {}

/** UI-language CJK fallback for the print stack (mirrors the :lang() variables in styles.css) */
function printCjkFonts(lang: Lang): string {
  switch (lang) {
    case 'ja':
      return "'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Yu Gothic UI', 'Yu Gothic', 'Meiryo'"
    case 'ko':
      return "'Apple SD Gothic Neo', 'Malgun Gothic'"
    case 'zh-TW':
      return "'PingFang TC', 'Microsoft JhengHei'"
    default:
      return "'PingFang SC'"
  }
}

const MAX_PRINT_CELLS = 50_000

/// The slice of the Univer facade the layout needs (structural, so the
/// caller passes the FWorksheet through a cast).
export interface PrintWorksheet {
  getLastRow(): number
  getLastColumn(): number
  getRowHeight(row: number): number
  getColumnWidth(column: number): number
  getMergedRanges(): {
    getRow(): number
    getColumn(): number
    getWidth(): number
    getHeight(): number
  }[]
  getRange(
    row: number,
    column: number,
    numRows: number,
    numColumns: number,
  ): {
    getDisplayValues(): string[][]
    getValues(): unknown[][]
  }
  getRange(row: number, column: number): { getCellStyleData(): PrintCellStyle | null }
}

/// The IStyleData fields the layout reads (all optional in Univer).
interface PrintCellStyle {
  readonly bl?: number
  readonly it?: number
  readonly ul?: { s?: number } | null
  readonly st?: { s?: number } | null
  readonly fs?: number
  readonly ff?: string | null
  readonly cl?: { rgb?: string | null } | null
  readonly bg?: { rgb?: string | null } | null
  readonly ht?: number
  readonly vt?: number
  readonly tb?: number
  readonly bd?: Partial<
    Record<'t' | 'b' | 'l' | 'r', { cl?: { rgb?: string | null } | null } | null>
  > | null
}

/// Inches, mirroring the gateway's Margins presets.
const MARGIN_PRESETS = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 },
  wide: { left: 1, right: 1, top: 1, bottom: 1 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75 },
} as const

/// OOXML paper-size code → Electron pageSize (custom sizes in inches).
const PAPER_SIZES: Record<number, WorkbookExportPdfRequest['pageSize']> = {
  1: 'Letter',
  3: 'Tabloid',
  5: 'Legal',
  7: { width: 7.25, height: 10.5 },
  8: 'A3',
  9: 'A4',
  11: 'A5',
}

const PAPER_WIDTH_INCHES: Record<string, number> = {
  Letter: 8.5,
  Tabloid: 11,
  Legal: 8.5,
  A3: 11.69,
  A4: 8.27,
  A5: 5.83,
}

export function buildSheetPrintPayload(
  worksheet: PrintWorksheet,
  pageSetup: PageSetupJournalState,
  fileName: string,
  sheetName: string,
): WorkbookExportPdfRequest {
  const area = pageSetup.printArea ? parseArea(pageSetup.printArea) : usedArea(worksheet)
  const rows = area.endRow - area.startRow + 1
  const columns = area.endColumn - area.startColumn + 1
  if (rows < 1 || columns < 1) throw new PrintError(t('appPrintNothing'))
  if (rows * columns > MAX_PRINT_CELLS) {
    throw new PrintError(t('appPrintTooLarge'))
  }

  const titles = pageSetup.printTitles ? parseTitleRows(pageSetup.printTitles) : null
  const grid = worksheet.getRange(area.startRow, area.startColumn, rows, columns)
  const display = grid.getDisplayValues()
  const raw = grid.getValues()
  const merges = mergeMaps(worksheet, area)

  const headings = pageSetup.printHeadings === true
  const gridlines = pageSetup.printGridlines === true
  const columnWidthsPt = Array.from(
    { length: columns },
    (_, offset) => worksheet.getColumnWidth(area.startColumn + offset) * 0.75,
  )
  const rowHeaderPt = headings ? 24 : 0

  const bodyRow = (row: number): string => {
    const cells: string[] = []
    if (headings) {
      cells.push(`<th class="hd">${row + 1}</th>`)
    }
    for (let column = area.startColumn; column <= area.endColumn; column += 1) {
      const key = `${row}:${column}`
      if (merges.covered.has(key)) continue
      const anchor = merges.anchors.get(key)
      const span = anchor
        ? ` rowspan="${Math.min(anchor.rows, area.endRow - row + 1)}"` +
          ` colspan="${Math.min(anchor.columns, area.endColumn - column + 1)}"`
        : ''
      const inArea = row >= area.startRow && row <= area.endRow
      const text = inArea
        ? (display[row - area.startRow]?.[column - area.startColumn] ?? '')
        : cellDisplay(worksheet, row, column)
      const rawValue = inArea ? raw[row - area.startRow]?.[column - area.startColumn] : undefined
      const style = worksheet.getRange(row, column).getCellStyleData()
      cells.push(
        `<td${span} style="${cellCss(style, rawValue, gridlines)}">${escapeHtml(text)}</td>`,
      )
    }
    const heightPt = Math.max(worksheet.getRowHeight(row) * 0.75, 10)
    return `<tr style="height:${round(heightPt)}pt">${cells.join('')}</tr>`
  }

  // Session header/footer ride the table's thead/tfoot, which Chromium
  // repeats at the top/bottom of every printed page.
  const totalColumns = columns + (headings ? 1 : 0)
  const resolveCodes = (text: string): string =>
    resolveHeaderFooterText(text, fileName.replace(/\.pdf$/, ''), sheetName, new Date())
  const headerRow = headerFooterRow(pageSetup.header, 'th', totalColumns, resolveCodes)
  const footerRow = headerFooterRow(pageSetup.footer, 'td', totalColumns, resolveCodes)

  const headParts: string[] = []
  if (headerRow !== '') headParts.push(headerRow)
  if (headings) {
    const letters = Array.from(
      { length: columns },
      (_, offset) => `<th class="hd">${columnLabel(area.startColumn + offset)}</th>`,
    )
    headParts.push(`<tr>${headings ? '<th class="hd"></th>' : ''}${letters.join('')}</tr>`)
  }
  if (titles) {
    for (let row = titles.start; row <= titles.end; row += 1) headParts.push(bodyRow(row))
  }

  const bodyParts: string[] = []
  for (let row = area.startRow; row <= area.endRow; row += 1) {
    // Title rows already repeat via the table header.
    if (titles && row >= titles.start && row <= titles.end) continue
    bodyParts.push(bodyRow(row))
  }

  const colgroup = `<colgroup>${headings ? `<col style="width:${rowHeaderPt}pt">` : ''}${columnWidthsPt
    .map((width) => `<col style="width:${round(width)}pt">`)
    .join('')}</colgroup>`
  const html =
    `<!doctype html><html lang="${htmlLang(getLang())}"><head><meta charset="utf-8"><style>
* { box-sizing: border-box; }
body { margin: 0; font-family: Calibri, 'Helvetica Neue', Arial, ${printCjkFonts(getLang())}, sans-serif; }
table { border-collapse: collapse; table-layout: fixed; }
thead { display: table-header-group; }
tfoot { display: table-footer-group; }
td, th { overflow: hidden; padding: 1pt 3pt; font-size: 11pt; vertical-align: bottom; }
th.hd { background: #f1f1f1; border: 0.5pt solid #b7b7b7; color: #444;
  font-size: 8.5pt; font-weight: 400; text-align: center; vertical-align: middle; }
th.hf, td.hf { padding: 2pt 0 6pt; font-weight: 400; }
td.hf { padding: 6pt 0 0; }
.hf > div { display: flex; font-size: 9pt; color: #333; }
.hf span { flex: 1; white-space: pre; }
.hf span:nth-child(2) { text-align: center; }
.hf span:last-child { text-align: right; }
</style></head><body><table>${colgroup}<thead>${headParts.join('')}</thead>` +
    `<tbody>${bodyParts.join('')}</tbody>` +
    `${footerRow === '' ? '' : `<tfoot>${footerRow}</tfoot>`}</table></body></html>`

  const margins = MARGIN_PRESETS[pageSetup.margins ?? 'normal']
  const pageSize = PAPER_SIZES[pageSetup.paperSize ?? 9] ?? 'A4'
  const landscape = pageSetup.orientation === 'landscape'
  return {
    fileName,
    html,
    landscape,
    pageSize,
    margins: { top: margins.top, bottom: margins.bottom, left: margins.left, right: margins.right },
    scale: computeScale(
      pageSetup,
      pageSize,
      landscape,
      margins,
      rowHeaderPt + columnWidthsPt.reduce((total, width) => total + width, 0),
    ),
  }
}

/// Resolves the field codes a static layout can know (&D date, &T time,
/// &F file name, &A sheet name, && literal &); page-dependent codes (&P,
/// &N) render empty — the HTML-to-PDF path has no page counter.
function resolveHeaderFooterText(
  text: string,
  fileName: string,
  sheetName: string,
  now: Date,
): string {
  return text.replace(/&(&|[A-Za-z])/g, (match, code: string) => {
    switch (code) {
      case '&':
        return '&'
      case 'P':
      case 'N':
        return ''
      case 'D':
        return now.toLocaleDateString()
      case 'T':
        return now.toLocaleTimeString()
      case 'F':
        return fileName
      case 'A':
        return sheetName
      default:
        return match
    }
  })
}

/// One left/center/right header or footer row spanning the whole table
/// ('' when the parts are empty).
function headerFooterRow(
  parts: HeaderFooterParts | null | undefined,
  cell: 'th' | 'td',
  colspan: number,
  resolveCodes: (text: string) => string,
): string {
  const sections = [parts?.left, parts?.center, parts?.right].map((text) =>
    resolveCodes(text ?? ''),
  )
  if (sections.every((text) => text === '')) return ''
  const spans = sections.map((text) => `<span>${escapeHtml(text)}</span>`).join('')
  return `<tr><${cell} class="hf" colspan="${colspan}"><div>${spans}</div></${cell}></tr>`
}

/// Excel's fit-to-width only shrinks; an explicit scale applies as-is.
function computeScale(
  pageSetup: PageSetupJournalState,
  pageSize: WorkbookExportPdfRequest['pageSize'],
  landscape: boolean,
  margins: { left: number; right: number; top: number; bottom: number },
  contentWidthPt: number,
): number {
  const fitPages = pageSetup.fitToPage === true ? (pageSetup.fitToWidth ?? 0) : 0
  if (fitPages > 0) {
    const paperWidthIn =
      typeof pageSize === 'string'
        ? landscape
          ? paperHeightInches(pageSize)
          : (PAPER_WIDTH_INCHES[pageSize] ?? 8.27)
        : landscape
          ? pageSize.height
          : pageSize.width
    const printableWidthPt = (paperWidthIn - margins.left - margins.right) * 72
    return clamp((printableWidthPt * fitPages) / contentWidthPt, 0.1, 1)
  }
  if (pageSetup.fitToPage !== true && pageSetup.scale !== undefined) {
    return clamp(pageSetup.scale / 100, 0.1, 2)
  }
  return 1
}

function paperHeightInches(name: string): number {
  const heights: Record<string, number> = {
    Letter: 11,
    Tabloid: 17,
    Legal: 14,
    A3: 16.54,
    A4: 11.69,
    A5: 8.27,
  }
  return heights[name] ?? 11.69
}

function usedArea(worksheet: PrintWorksheet) {
  return {
    startRow: 0,
    startColumn: 0,
    endRow: Math.max(worksheet.getLastRow(), 0),
    endColumn: Math.max(worksheet.getLastColumn(), 0),
  }
}

function parseArea(area: string) {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(area)
  if (!match) throw new PrintError(t('appPrintBadArea', { area }))
  return {
    startRow: Number(match[2]) - 1,
    startColumn: columnIndex(match[1] ?? 'A'),
    endRow: Number(match[4]) - 1,
    endColumn: columnIndex(match[3] ?? 'A'),
  }
}

function parseTitleRows(titles: string): { start: number; end: number } {
  const match = /^(\d{1,7}):(\d{1,7})$/.exec(titles)
  if (!match) throw new PrintError(t('appPrintBadTitles', { titles }))
  const start = Number(match[1]) - 1
  const end = Number(match[2]) - 1
  if (end - start > 20) throw new PrintError(t('appPrintTitlesLimit'))
  return { start, end }
}

function mergeMaps(
  worksheet: PrintWorksheet,
  area: { startRow: number; endRow: number; startColumn: number; endColumn: number },
) {
  const anchors = new Map<string, { rows: number; columns: number }>()
  const covered = new Set<string>()
  for (const merge of worksheet.getMergedRanges()) {
    const row = merge.getRow()
    const column = merge.getColumn()
    if (row > area.endRow || column > area.endColumn) continue
    if (row + merge.getHeight() - 1 < area.startRow) continue
    if (column + merge.getWidth() - 1 < area.startColumn) continue
    anchors.set(`${row}:${column}`, { rows: merge.getHeight(), columns: merge.getWidth() })
    for (let r = row; r < row + merge.getHeight(); r += 1) {
      for (let c = column; c < column + merge.getWidth(); c += 1) {
        if (r !== row || c !== column) covered.add(`${r}:${c}`)
      }
    }
  }
  return { anchors, covered }
}

function cellDisplay(worksheet: PrintWorksheet, row: number, column: number): string {
  return worksheet.getRange(row, column, 1, 1).getDisplayValues()[0]?.[0] ?? ''
}

function cellCss(style: PrintCellStyle | null, rawValue: unknown, gridlines: boolean): string {
  const rules: string[] = []
  if (style?.bl === 1) rules.push('font-weight:700')
  if (style?.it === 1) rules.push('font-style:italic')
  const decorations = [
    style?.ul?.s === 1 ? 'underline' : '',
    style?.st?.s === 1 ? 'line-through' : '',
  ].filter(Boolean)
  if (decorations.length > 0) rules.push(`text-decoration:${decorations.join(' ')}`)
  if (style?.fs) rules.push(`font-size:${round(style.fs)}pt`)
  // a font name comes straight from styles.xml; anything outside the whitelist could
  // close the style attribute and inject markup into the exported page
  const family = style?.ff?.replace(/[^\p{L}\p{N} \-_.]/gu, '')
  // fallbacks mirror the body stack: an uninstalled family (e.g. Aptos)
  // must not drop to the browser's serif default in the exported page
  if (family)
    rules.push(
      `font-family:'${family}',Calibri,'Helvetica Neue',Arial,${printCjkFonts(getLang())},sans-serif`,
    )
  if (style?.cl?.rgb) rules.push(`color:${cssColor(style.cl.rgb)}`)
  if (style?.bg?.rgb) rules.push(`background:${cssColor(style.bg.rgb)}`)
  const align =
    style?.ht === 1
      ? 'left'
      : style?.ht === 2
        ? 'center'
        : style?.ht === 3
          ? 'right'
          : typeof rawValue === 'number'
            ? 'right'
            : typeof rawValue === 'boolean'
              ? 'center'
              : 'left'
  rules.push(`text-align:${align}`)
  if (style?.vt === 1) rules.push('vertical-align:top')
  else if (style?.vt === 2) rules.push('vertical-align:middle')
  rules.push(style?.tb === 3 ? 'white-space:pre-wrap;word-break:break-word' : 'white-space:pre')
  const defaultBorder = gridlines ? '0.5pt solid #c0c0c0' : 'none'
  for (const [edge, css] of [
    ['t', 'top'],
    ['b', 'bottom'],
    ['l', 'left'],
    ['r', 'right'],
  ]) {
    const border = style?.bd?.[edge as 't' | 'b' | 'l' | 'r']
    rules.push(
      `border-${css}:${
        border ? `0.75pt solid ${cssColor(border.cl?.rgb ?? '#000000')}` : defaultBorder
      }`,
    )
  }
  return rules.join(';')
}

function cssColor(rgb: string): string {
  return /^(#[0-9a-fA-F]{3,8}|rgba?\([\d ,.%]+\))$/.test(rgb) ? rgb : '#000'
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

function round(value: number): number {
  return Math.round(value * 100) / 100
}
