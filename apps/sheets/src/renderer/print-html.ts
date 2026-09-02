/// Lays the active sheet out as print HTML from the live Univer model —
/// display strings (number formats applied), cell styles, merges, and the
/// sheet's effective page setup (print areas, repeated title rows, gridlines,
/// headings, header/footer). The main process turns the HTML into a PDF.

import { BorderStyleTypes } from '@univerjs/core'
import { htmlLang, type Lang } from '@genoffice/i18n'
import { columnIndex, columnLabel } from '../domain/cell-address'

import type { WorkbookExportPdfRequest } from '../shared/desktop-api'
import type { HeaderFooterParts } from './edit-journal'
import {
  fitToPageScale,
  MAX_PRINT_SCALE,
  MIN_PRINT_SCALE,
  type PrintAreaHeights,
} from './print-scale'
import type { EffectivePageSetup, HeaderFooterPair, PrintMargins } from './print-settings'
import { getLang, t } from './i18n/locale'

export class PrintError extends Error {}

/// A `&G` picture resolved to bytes for the print templates.
export interface HeaderFooterPictureImage {
  readonly dataUrl: string
  readonly widthPt: number
  readonly heightPt: number
}

/// Pictures keyed by VML slot: L/C/R × H/F plus an EVEN or FIRST suffix for
/// the page variants (`LH`, `CFFIRST`, `RHEVEN`).
export type HeaderFooterPictures = ReadonlyMap<string, HeaderFooterPictureImage>

/// Pictures for the three sections of one header or footer.
export interface SectionPictures {
  readonly left?: HeaderFooterPictureImage | undefined
  readonly center?: HeaderFooterPictureImage | undefined
  readonly right?: HeaderFooterPictureImage | undefined
}

type PageVariant = 'odd' | 'even' | 'first'

/// Chromium lays the print body out with Calibri 11pt unless the cell says
/// otherwise; a text row is at least one line plus the cell padding tall,
/// which can exceed Excel's saved row height by a point or so — the
/// fit-to-page pagination must count the printed height, not the saved one.
const LINE_HEIGHT_FACTOR = 1.25
const CELL_VERTICAL_PADDING_PT = 2
const DEFAULT_FONT_SIZE_PT = 11
/// The row/column heading strip (8.5pt text, padding, border).
const HEADING_ROW_HEIGHT_PT = 14

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
    Record<'t' | 'b' | 'l' | 'r', { s?: number; cl?: { rgb?: string | null } | null } | null>
  > | null
}

/// Print weight of a cell border by Univer BorderStyleTypes value: Excel
/// prints thin at 0.75pt (1px @ 96dpi), medium at 1.5pt and thick at
/// 2.25pt — the same 1 : 2 : 3 ladder the grid draws. Dash patterns keep
/// printing solid (unchanged); only the weight is mapped here.
export function printBorderWidthPt(style: number | undefined): number {
  switch (style) {
    case BorderStyleTypes.MEDIUM:
    case BorderStyleTypes.MEDIUM_DASHED:
    case BorderStyleTypes.MEDIUM_DASH_DOT:
    case BorderStyleTypes.MEDIUM_DASH_DOT_DOT:
      return 1.5
    case BorderStyleTypes.THICK:
      return 2.25
    default:
      return 0.75
  }
}

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
  setup: EffectivePageSetup,
  fileName: string,
  sheetName: string,
  pictures: HeaderFooterPictures = new Map(),
): WorkbookExportPdfRequest {
  const areas =
    setup.printAreas.length > 0 ? setup.printAreas.map(parseArea) : [usedArea(worksheet)]
  const titles = setup.printTitles ? parseTitleRows(setup.printTitles) : null
  const headings = setup.printHeadings
  const gridlines = setup.printGridlines
  const rowHeaderPt = headings ? 24 : 0

  let totalCells = 0
  for (const area of areas) {
    const rows = area.endRow - area.startRow + 1
    const columns = area.endColumn - area.startColumn + 1
    if (rows < 1 || columns < 1) throw new PrintError(t('appPrintNothing'))
    totalCells += rows * columns
  }
  if (totalCells > MAX_PRINT_CELLS) throw new PrintError(t('appPrintTooLarge'))

  let maxContentWidthPt = 0
  const tables: string[] = []
  const areaHeights: PrintAreaHeights[] = []
  for (const area of areas) {
    const rows = area.endRow - area.startRow + 1
    const columns = area.endColumn - area.startColumn + 1
    const grid = worksheet.getRange(area.startRow, area.startColumn, rows, columns)
    const display = grid.getDisplayValues()
    const raw = grid.getValues()
    const merges = mergeMaps(worksheet, area)
    const columnWidthsPt = Array.from(
      { length: columns },
      (_, offset) => worksheet.getColumnWidth(area.startColumn + offset) * 0.75,
    )
    maxContentWidthPt = Math.max(
      maxContentWidthPt,
      rowHeaderPt + columnWidthsPt.reduce((total, width) => total + width, 0),
    )

    // Printed height of the row just laid out by bodyRow (saved height, or
    // taller when a cell's text line does not fit it).
    let printedRowHeightPt = 0
    const bodyRow = (row: number): string => {
      const cells: string[] = []
      let textHeightPt = 0
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
        if (text !== '' && !anchor) {
          textHeightPt = Math.max(
            textHeightPt,
            (style?.fs ?? DEFAULT_FONT_SIZE_PT) * LINE_HEIGHT_FACTOR + CELL_VERTICAL_PADDING_PT,
          )
        }
        cells.push(
          `<td${span} style="${cellCss(style, rawValue, gridlines)}">${escapeHtml(text)}</td>`,
        )
      }
      const heightPt = Math.max(worksheet.getRowHeight(row) * 0.75, 10)
      printedRowHeightPt = Math.max(heightPt, textHeightPt)
      return `<tr style="height:${round(heightPt)}pt">${cells.join('')}</tr>`
    }

    const headParts: string[] = []
    let repeatedHeightPt = headings ? HEADING_ROW_HEIGHT_PT : 0
    if (headings) {
      const letters = Array.from(
        { length: columns },
        (_, offset) => `<th class="hd">${columnLabel(area.startColumn + offset)}</th>`,
      )
      headParts.push(`<tr><th class="hd"></th>${letters.join('')}</tr>`)
    }
    if (titles) {
      for (let row = titles.start; row <= titles.end; row += 1) {
        headParts.push(bodyRow(row))
        repeatedHeightPt += printedRowHeightPt
      }
    }

    const bodyParts: string[] = []
    const rowHeightsPt: number[] = []
    for (let row = area.startRow; row <= area.endRow; row += 1) {
      // Title rows already repeat via the table header.
      if (titles && row >= titles.start && row <= titles.end) continue
      bodyParts.push(bodyRow(row))
      rowHeightsPt.push(printedRowHeightPt)
    }
    areaHeights.push({ repeatedHeightPt, rowHeightsPt })

    const colgroup = `<colgroup>${headings ? `<col style="width:${rowHeaderPt}pt">` : ''}${columnWidthsPt
      .map((width) => `<col style="width:${round(width)}pt">`)
      .join('')}</colgroup>`
    tables.push(
      `<table>${colgroup}<thead>${headParts.join('')}</thead><tbody>${bodyParts.join('')}</tbody></table>`,
    )
  }

  const html =
    `<!doctype html><html lang="${htmlLang(getLang())}"><head><meta charset="utf-8"><style>
* { box-sizing: border-box; }
body { margin: 0; font-family: Calibri, 'Helvetica Neue', Arial, ${printCjkFonts(getLang())}, sans-serif; }
table { border-collapse: collapse; table-layout: fixed; }
table + table { break-before: page; }
thead { display: table-header-group; }
td, th { overflow: hidden; padding: 1pt 3pt; font-size: 11pt; vertical-align: bottom; }
th.hd { background: #f1f1f1; border: 0.5pt solid #b7b7b7; color: #444;
  font-size: 8.5pt; font-weight: 400; text-align: center; vertical-align: middle; }
</style></head><body>` +
    tables.join('') +
    `</body></html>`

  const margins = setup.margins
  const pageSize = PAPER_SIZES[setup.paperSize] ?? 'A4'
  const landscape = setup.orientation === 'landscape'
  const now = new Date()
  const baseName = fileName.replace(/\.pdf$/, '')
  const scale = computeScale(setup, pageSize, landscape, margins, maxContentWidthPt, areaHeights)
  // Excel's "scale with document" (the default) shrinks the header/footer
  // text and pictures by the same factor as the sheet.
  const templateScale = setup.headerFooterScaleWithDoc ? scale : 1
  const templates = (pair: HeaderFooterPair, variant: PageVariant) => {
    const headerTemplate = pair.header
      ? buildHeaderFooterTemplate(
          pair.header,
          'header',
          margins,
          baseName,
          sheetName,
          now,
          sectionPictures(pictures, 'header', variant),
          templateScale,
        )
      : undefined
    const footerTemplate = pair.footer
      ? buildHeaderFooterTemplate(
          pair.footer,
          'footer',
          margins,
          baseName,
          sheetName,
          now,
          sectionPictures(pictures, 'footer', variant),
          templateScale,
        )
      : undefined
    return {
      ...(headerTemplate === undefined ? {} : { headerTemplate }),
      ...(footerTemplate === undefined ? {} : { footerTemplate }),
    }
  }
  return {
    fileName,
    html,
    landscape,
    pageSize,
    margins: { top: margins.top, bottom: margins.bottom, left: margins.left, right: margins.right },
    scale,
    ...templates({ header: setup.header, footer: setup.footer }, 'odd'),
    ...(setup.firstPage === null ? {} : { firstPage: templates(setup.firstPage, 'first') }),
    ...(setup.evenPages === null ? {} : { evenPages: templates(setup.evenPages, 'even') }),
  }
}

/// The `&G` pictures of one header or footer's three sections, for one page
/// variant (VML slot ids: LH/CH/RH, LF/CF/RF, plus EVEN/FIRST).
export function sectionPictures(
  pictures: HeaderFooterPictures,
  kind: 'header' | 'footer',
  variant: PageVariant,
): SectionPictures {
  const suffix = variant === 'odd' ? '' : variant.toUpperCase()
  const slot = (section: 'L' | 'C' | 'R') =>
    pictures.get(`${section}${kind === 'header' ? 'H' : 'F'}${suffix}`)
  const left = slot('L')
  const center = slot('C')
  const right = slot('R')
  return {
    ...(left === undefined ? {} : { left }),
    ...(center === undefined ? {} : { center }),
    ...(right === undefined ? {} : { right }),
  }
}

/// Header/footer text size before scaleWithDoc applies.
const HEADER_FOOTER_FONT_SIZE_PT = 9

/// One left/center/right header or footer as a Chromium print template
/// (rendered in the page's margin box; undefined when the parts are empty).
/// `scale` is the print scale the text and pictures follow (1 when the
/// header/footer keeps its size).
export function buildHeaderFooterTemplate(
  parts: HeaderFooterParts,
  kind: 'header' | 'footer',
  margins: PrintMargins,
  fileName: string,
  sheetName: string,
  now: Date,
  pictures: SectionPictures = {},
  scale = 1,
): string | undefined {
  const sections = [parts.left ?? '', parts.center ?? '', parts.right ?? '']
  if (sections.every((text) => text === '')) return undefined
  const sectionPicture = [pictures.left, pictures.center, pictures.right]
  const rendered = sections.map((text, index) =>
    renderHeaderFooterHtml(text, fileName, sheetName, now, sectionPicture[index], scale),
  )
  const fontSizePt = round(HEADER_FOOTER_FONT_SIZE_PT * scale)
  // Excel offsets the header/footer from the paper edge by its own margin.
  const offset =
    kind === 'header'
      ? `padding-top:${round(margins.header)}in`
      : `padding-bottom:${round(margins.footer)}in`
  // Equal thirds like Excel's sections; an oversized picture or unbreakable
  // text overflows its neighbours instead of squeezing them.
  const spanStyle = 'flex:1;min-width:0;white-space:pre-wrap'
  // Chromium's template document is content-box; without an inline
  // border-box the width:100% + side padding overflows the page and
  // shifts/clips the sections.
  return (
    `<div style="box-sizing:border-box;display:flex;width:100%;font-size:${fontSizePt}pt;color:#000;` +
    `font-family:Calibri,'Helvetica Neue',Arial,sans-serif;` +
    `padding-left:${round(margins.left)}in;padding-right:${round(margins.right)}in;${offset}">` +
    `<span style="${spanStyle}">${rendered[0]}</span>` +
    `<span style="${spanStyle};text-align:center">${rendered[1]}</span>` +
    `<span style="${spanStyle};text-align:right">${rendered[2]}</span></div>`
  )
}

/// Field codes → template HTML: &P/&N become Chromium's live pageNumber/
/// totalPages spans, static codes (&D &T &F &A, && literal) resolve now,
/// &G becomes the section's picture (nothing when the slot has none, like
/// Excel), everything else is HTML-escaped verbatim.
export function renderHeaderFooterHtml(
  text: string,
  fileName: string,
  sheetName: string,
  now: Date,
  picture?: HeaderFooterPictureImage,
  scale = 1,
): string {
  let html = ''
  let literal = ''
  const flush = (): void => {
    html += escapeHtml(literal)
    literal = ''
  }
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index] ?? ''
    if (character !== '&') {
      literal += character
      continue
    }
    const code = text[index + 1]
    if (code === undefined) {
      literal += '&'
      break
    }
    index += 1
    switch (code) {
      case '&':
        literal += '&'
        break
      case 'P':
        flush()
        html += '<span class="pageNumber"></span>'
        break
      case 'N':
        flush()
        html += '<span class="totalPages"></span>'
        break
      case 'D':
        literal += now.toLocaleDateString()
        break
      case 'T':
        literal += now.toLocaleTimeString()
        break
      case 'F':
        literal += fileName
        break
      case 'A':
        literal += sheetName
        break
      case 'G':
        if (picture) {
          flush()
          html += pictureHtml(picture, scale)
        }
        break
      default:
        literal += `&${code}`
    }
  }
  flush()
  return html
}

/// The picture at its declared size times the print scale (points → CSS px
/// at 96/72). The data URL is built from a validated media type and base64
/// payload; escaping it anyway keeps the attribute closed no matter what.
function pictureHtml(picture: HeaderFooterPictureImage, scale: number): string {
  const width = round((picture.widthPt * scale * 96) / 72)
  const height = round((picture.heightPt * scale * 96) / 72)
  return (
    `<img src="${escapeAttribute(picture.dataUrl)}" ` +
    `style="width:${width}px;height:${height}px;vertical-align:bottom">`
  )
}

/// Excel's fit-to-page only shrinks; an explicit scale applies as-is.
function computeScale(
  setup: EffectivePageSetup,
  pageSize: WorkbookExportPdfRequest['pageSize'],
  landscape: boolean,
  margins: { left: number; right: number; top: number; bottom: number },
  contentWidthPt: number,
  areas: readonly PrintAreaHeights[],
): number {
  if (!setup.fitToPage) {
    return clamp(setup.scale / 100, MIN_PRINT_SCALE, MAX_PRINT_SCALE)
  }
  const [paperWidthIn, paperHeightIn] =
    typeof pageSize === 'string'
      ? [PAPER_WIDTH_INCHES[pageSize] ?? 8.27, paperHeightInches(pageSize)]
      : [pageSize.width, pageSize.height]
  const [acrossIn, downIn] = landscape
    ? [paperHeightIn, paperWidthIn]
    : [paperWidthIn, paperHeightIn]
  return fitToPageScale({
    printableWidthPt: (acrossIn - margins.left - margins.right) * 72,
    printableHeightPt: (downIn - margins.top - margins.bottom) * 72,
    fitToWidth: setup.fitToWidth,
    fitToHeight: setup.fitToHeight,
    contentWidthPt,
    areas,
  })
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
        border
          ? `${printBorderWidthPt(border.s)}pt solid ${cssColor(border.cl?.rgb ?? '#000000')}`
          : defaultBorder
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

function escapeAttribute(value: string): string {
  return escapeHtml(value).replace(/"/g, '&quot;')
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

function round(value: number): number {
  return Math.round(value * 100) / 100
}
