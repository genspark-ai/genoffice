/// Lays a sheet out as printed pages from the live Univer model — display
/// strings (number formats applied), cell styles, merges, and the sheet's
/// effective page setup (print areas, repeated titles, manual breaks,
/// gridlines, headings, header/footer). Every page is an explicitly sized
/// paper box with its header/footer inside, so the same HTML serves the
/// preview, the system print and the PDF export, and a job may carry
/// several sheets with different paper or orientation.

import { BorderStyleTypes } from '@univerjs/core'
import { htmlLang, type Lang } from '@genoffice/i18n'
import { columnIndex, columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'

import type { WorkbookExportPdfRequest } from '../shared/desktop-api'
import type { HeaderFooterParts } from './edit-journal'
import { expandHeaderFooterCodes } from './print-hf-codes'
import {
  bodyColumns,
  bodyRows,
  planPrintPages,
  type IndexSpan,
  type PaginationInput,
  type PrintArea,
  type PrintPage,
} from './print-paginate'
import {
  fitToPageScale,
  MAX_PRINT_SCALE,
  MIN_PRINT_SCALE,
  type PrintAreaHeights,
} from './print-scale'
import type { EffectivePageSetup, HeaderFooterPair, PrintMargins } from './print-settings'
import type { PrintVisual, PrintVisualSnapshot } from './print-visuals'
import { getLang, t } from './i18n/locale'

export class PrintError extends Error {}

/// A `&G` picture resolved to bytes for the print pages.
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

/// Excel prints gridlines as hairlines.
const GRIDLINE_PT = 0.75
/// The row/column heading strip (8.5pt text, padding, border).
const HEADING_ROW_HEIGHT_PT = 14
const HEADING_COLUMN_WIDTH_PT = 24
/// Header/footer text size before scaleWithDoc applies.
const HEADER_FOOTER_FONT_SIZE_PT = 9

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

/**
 * Univer dimension boundary: a corrupt workbook can report NaN/negative
 * widths/heights, which previously flowed into colgroup styles, left/top
 * accumulators, and scale math as NaNpt. Clamp to a finite positive value.
 */
function finitePt(value: number, fallback: number): number {
  if (!Number.isFinite(value) || value < 0) return fallback
  return Math.min(value, 100000)
}

/// The slice of the Univer facade the layout needs (structural, so the
/// caller passes the FWorksheet through a cast).
export interface PrintWorksheet {
  getLastRow(): number
  getLastColumn(): number
  getRowHeight(row: number): number
  getColumnWidth(column: number): number
  /// Hidden rows/columns do not print (Excel); absent in older fakes.
  getSheet?(): { getRowVisible(row: number): boolean; getColVisible(column: number): boolean }
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
/// ECMA-376 §18.3.1.70: unmapped codes previously fell back to A4 and
/// mis-scaled B4/B5/Folio/Statement output.
const PAPER_SIZES: Record<number, WorkbookExportPdfRequest['pageSize']> = {
  1: 'Letter',
  2: 'Letter',
  3: 'Tabloid',
  4: { width: 17, height: 11 },
  5: 'Legal',
  6: { width: 5.5, height: 8.5 },
  7: { width: 7.25, height: 10.5 },
  8: 'A3',
  9: 'A4',
  10: 'A4',
  11: 'A5',
  12: { width: 9.84, height: 13.9 },
  13: { width: 7.17, height: 10.12 },
  14: { width: 8.5, height: 13 },
  15: { width: 8.46, height: 10.83 },
  16: { width: 10, height: 14 },
  18: 'Letter',
}

const NAMED_PAPER_INCHES: Record<string, readonly [number, number]> = {
  Letter: [8.5, 11],
  Tabloid: [11, 17],
  Legal: [8.5, 14],
  A3: [11.69, 16.54],
  A4: [8.27, 11.69],
  A5: [5.83, 8.27],
}

export interface PaperGeometry {
  readonly pageSize: WorkbookExportPdfRequest['pageSize']
  readonly landscape: boolean
  /// Oriented paper size.
  readonly widthIn: number
  readonly heightIn: number
}

export function paperGeometry(
  paperSize: number,
  orientation: 'portrait' | 'landscape',
): PaperGeometry {
  const pageSize = PAPER_SIZES[paperSize] ?? 'A4'
  const [portraitWidth, portraitHeight] =
    typeof pageSize === 'string'
      ? (NAMED_PAPER_INCHES[pageSize] ?? NAMED_PAPER_INCHES.A4 ?? [8.27, 11.69])
      : [pageSize.width, pageSize.height]
  const landscape = orientation === 'landscape'
  return {
    pageSize,
    landscape,
    widthIn: landscape ? portraitHeight : portraitWidth,
    heightIn: landscape ? portraitWidth : portraitHeight,
  }
}

export interface SheetPrintOptions {
  /// Manual page breaks (0-based index of the row/column that starts a page).
  readonly breaks?: { rowBreaks: readonly number[]; colBreaks: readonly number[] } | null
  /// Print this A1 range instead of the sheet's print area (Excel's "Print Selection").
  readonly selection?: string | null
  readonly now?: Date
}

/// One laid-out page before the job-wide numbering is known.
export interface SheetPrintPageRecord {
  readonly body: string
  readonly page: PrintPage
}

/// The sheet's header/footer texts by page variant (null variant = off).
export interface SheetHeaderFooter {
  readonly odd: HeaderFooterPair
  readonly first: HeaderFooterPair | null
  readonly even: HeaderFooterPair | null
}

export interface SheetPrintDocument {
  readonly sheetName: string
  readonly fileName: string
  readonly paper: PaperGeometry
  readonly margins: PrintMargins
  /// Content scale actually applied (fit-to-page resolved).
  readonly scale: number
  readonly headerFooterScale: number
  readonly pictures: HeaderFooterPictures
  readonly headerFooter: SheetHeaderFooter
  readonly css: string
  readonly pages: readonly SheetPrintPageRecord[]
  readonly now: Date
}

export function layoutSheetPrint(
  worksheet: PrintWorksheet,
  setup: EffectivePageSetup,
  fileName: string,
  sheetName: string,
  pictures: HeaderFooterPictures = new Map(),
  visuals: PrintVisualSnapshot = { visuals: [], css: '' },
  options: SheetPrintOptions = {},
): SheetPrintDocument {
  const areas =
    options.selection !== undefined && options.selection !== null
      ? [parseArea(options.selection)]
      : setup.printAreas.length > 0
        ? setup.printAreas.map(parseArea)
        : [usedArea(worksheet, visuals.visuals)]
  const titleRows = setup.printTitles ? parseTitleRows(setup.printTitles) : null
  const titleColumns = setup.printTitleColumns ? parseTitleColumns(setup.printTitleColumns) : null
  const headings = setup.printHeadings
  const gridlines = setup.printGridlines

  let totalCells = 0
  for (const area of areas) {
    const rows = area.endRow - area.startRow + 1
    const columns = area.endColumn - area.startColumn + 1
    if (rows < 1 || columns < 1) throw new PrintError(t('appPrintNothing'))
    totalCells += rows * columns
  }
  if (totalCells > MAX_PRINT_CELLS) throw new PrintError(t('appPrintTooLarge'))

  const sheet = worksheet.getSheet?.()
  const rowHeightPt = (row: number): number =>
    sheet && !sheet.getRowVisible(row) ? 0 : finitePt(worksheet.getRowHeight(row), 20) * 0.75
  const columnWidthPt = (column: number): number =>
    sheet && !sheet.getColVisible(column)
      ? 0
      : finitePt(worksheet.getColumnWidth(column), 64) * 0.75
  const headingWidthPt = headings ? HEADING_COLUMN_WIDTH_PT : 0
  const headingHeightPt = headings ? HEADING_ROW_HEIGHT_PT : 0
  const titleRowsHeightPt = spanTotal(titleRows, rowHeightPt)
  const titleColumnsWidthPt = spanTotal(titleColumns, columnWidthPt)

  const paper = paperGeometry(setup.paperSize, setup.orientation)
  const margins = setup.margins
  const printableWidthPt = Math.max((paper.widthIn - margins.left - margins.right) * 72, 1)
  const printableHeightPt = Math.max((paper.heightIn - margins.top - margins.bottom) * 72, 1)

  const baseInput = {
    areas,
    rowHeightPt,
    columnWidthPt,
    titleRows,
    titleColumns,
    headingWidthPt,
    headingHeightPt,
  }
  const areaHeights: PrintAreaHeights[] = []
  let contentWidthPt = 0
  for (const area of areas) {
    const rows = bodyRows(area, baseInput)
    const columns = bodyColumns(area, baseInput)
    areaHeights.push({
      repeatedHeightPt: headingHeightPt + titleRowsHeightPt,
      rowHeightsPt: rows.map(rowHeightPt),
    })
    contentWidthPt = Math.max(
      contentWidthPt,
      headingWidthPt + titleColumnsWidthPt + columns.reduce((sum, c) => sum + columnWidthPt(c), 0),
    )
  }
  const scale = setup.fitToPage
    ? fitToPageScale({
        printableWidthPt,
        printableHeightPt,
        fitToWidth: setup.fitToWidth,
        fitToHeight: setup.fitToHeight,
        contentWidthPt,
        areas: areaHeights,
      })
    : clamp(setup.scale / 100, MIN_PRINT_SCALE, MAX_PRINT_SCALE)

  // Excel ignores manual breaks while fit-to-page decides the scale.
  const breaks = setup.fitToPage ? null : options.breaks
  const plan: PaginationInput = {
    ...baseInput,
    rowBreaks: breaks?.rowBreaks ?? [],
    colBreaks: breaks?.colBreaks ?? [],
    pageWidthPt: printableWidthPt / scale,
    pageHeightPt: printableHeightPt / scale,
  }
  const pages = planPrintPages(plan)

  const areaData = areas.map((area) => {
    const rows = area.endRow - area.startRow + 1
    const columns = area.endColumn - area.startColumn + 1
    const grid = worksheet.getRange(area.startRow, area.startColumn, rows, columns)
    return {
      area,
      display: grid.getDisplayValues(),
      raw: grid.getValues(),
      // Repeated titles print with their merges even when they sit outside
      // the print area.
      merges: mergeMaps(worksheet, {
        startRow: Math.min(area.startRow, titleRows?.start ?? area.startRow),
        endRow: Math.max(area.endRow, titleRows?.end ?? area.endRow),
        startColumn: Math.min(area.startColumn, titleColumns?.start ?? area.startColumn),
        endColumn: Math.max(area.endColumn, titleColumns?.end ?? area.endColumn),
      }),
      rowOffset: offsets(bodyRows(area, plan), rowHeightPt),
      columnOffset: offsets(bodyColumns(area, plan), columnWidthPt),
    }
  })

  const displayAt = (data: (typeof areaData)[number], row: number, column: number): string => {
    const { area } = data
    if (
      row >= area.startRow &&
      row <= area.endRow &&
      column >= area.startColumn &&
      column <= area.endColumn
    ) {
      return data.display[row - area.startRow]?.[column - area.startColumn] ?? ''
    }
    return cellDisplay(worksheet, row, column)
  }
  const rawAt = (data: (typeof areaData)[number], row: number, column: number): unknown => {
    const { area } = data
    if (
      row >= area.startRow &&
      row <= area.endRow &&
      column >= area.startColumn &&
      column <= area.endColumn
    ) {
      return data.raw[row - area.startRow]?.[column - area.startColumn]
    }
    return undefined
  }

  const styleAt = (row: number, column: number): PrintCellStyle | null =>
    worksheet.getRange(row, column).getCellStyleData()
  /// Each shared edge has one owner: the upper/left cell paints it unless
  /// only the lower/right neighbour declares a border there. Page edges are
  /// painted by the cell at the edge; the heading strip supplies the outer
  /// top/left of the first row/column.
  interface Neighbours {
    readonly above: number | null
    readonly below: number | null
    readonly leftOf: number | null
    readonly rightOf: number | null
  }
  const cell = (
    data: (typeof areaData)[number],
    row: number,
    column: number,
    span: string,
    heightPt: number,
    near: Neighbours,
  ): string => {
    const style = styleAt(row, column)
    const text = displayAt(data, row, column)
    const align = style?.vt === 1 ? 'flex-start' : style?.vt === 2 ? 'center' : 'flex-end'
    const custom = (r: number | null, c: number | null, edge: 't' | 'b' | 'l' | 'r'): boolean =>
      r !== null && c !== null && styleAt(r, c)?.bd?.[edge] != null
    const edges = {
      top:
        style?.bd?.t != null
          ? near.above === null || !custom(near.above, column, 'b')
          : gridlines && near.above === null && !headings,
      left:
        style?.bd?.l != null
          ? near.leftOf === null || !custom(row, near.leftOf, 'r')
          : gridlines && near.leftOf === null && !headings,
      bottom: style?.bd?.b != null || (gridlines && !custom(near.below, column, 't')),
      right: style?.bd?.r != null || (gridlines && !custom(row, near.rightOf, 'l')),
    }
    return (
      `<td${span} style="${cellCss(style, rawAt(data, row, column), gridlines, edges)}">` +
      `<div class="c" style="height:${round(heightPt)}pt;align-items:${align}"><span>${escapeHtml(text)}</span></div></td>`
    )
  }

  const records: SheetPrintPageRecord[] = pages.map((page) => {
    const data = areaData[page.area]
    if (!data) throw new PrintError(t('appPrintNothing'))
    const pageRows = bodyRows({ ...data.area, startRow: page.rowStart, endRow: page.rowEnd }, plan)
    const pageColumns = bodyColumns(
      { ...data.area, startColumn: page.colStart, endColumn: page.colEnd },
      plan,
    )
    const titleRowList = titleRows ? spanIndices(titleRows).filter((r) => rowHeightPt(r) > 0) : []
    const titleColumnList = titleColumns
      ? spanIndices(titleColumns).filter((c) => columnWidthPt(c) > 0)
      : []
    const allColumns = [...titleColumnList, ...pageColumns]
    const printedColumns = new Set(allColumns)
    const printedRows = [...titleRowList, ...pageRows]

    const colgroup =
      `<colgroup>${headings ? `<col style="width:${HEADING_COLUMN_WIDTH_PT}pt">` : ''}` +
      allColumns.map((c) => `<col style="width:${round(columnWidthPt(c))}pt">`).join('') +
      `</colgroup>`

    // Merges span within their table section: title rows stay in the head,
    // body rows in the body; columns span every printed column.
    const rowHtml = (row: number, sectionRows: readonly number[], position: number): string => {
      const heightPt = rowHeightPt(row)
      const cells: string[] = []
      if (headings) cells.push(`<th class="hd hd-left">${row + 1}</th>`)
      const inSection = new Set(sectionRows)
      const above = printedRows[position - 1] ?? null
      const rowAfter = (last: number): number | null =>
        printedRows[printedRows.indexOf(last) + 1] ?? null
      const columnAfter = (last: number): number | null =>
        allColumns[allColumns.indexOf(last) + 1] ?? null
      allColumns.forEach((column, index) => {
        const key = `${row}:${column}`
        const near = {
          above,
          below: rowAfter(row),
          leftOf: allColumns[index - 1] ?? null,
          rightOf: columnAfter(column),
        }
        const coveredBy = data.merges.covered.get(key)
        if (coveredBy !== undefined) {
          const [anchorRow, anchorColumn] = coveredBy
          if (inSection.has(anchorRow) && printedColumns.has(anchorColumn)) return
          cells.push(cell(data, row, column, '', heightPt, near))
          return
        }
        const anchor = data.merges.anchors.get(key)
        if (anchor) {
          const spannedRows = sectionRows.filter((r) => r >= row && r < row + anchor.rows)
          const spannedColumns = allColumns.filter(
            (c) => c >= column && c < column + anchor.columns,
          )
          const spanned = spannedRows.reduce((sum, r) => sum + rowHeightPt(r), 0)
          const lastRow = spannedRows[spannedRows.length - 1] ?? row
          const lastColumn = spannedColumns[spannedColumns.length - 1] ?? column
          cells.push(
            cell(
              data,
              row,
              column,
              ` rowspan="${spannedRows.length}" colspan="${spannedColumns.length}"`,
              spanned,
              { ...near, below: rowAfter(lastRow), rightOf: columnAfter(lastColumn) },
            ),
          )
          return
        }
        cells.push(cell(data, row, column, '', heightPt, near))
      })
      return `<tr style="height:${round(heightPt)}pt">${cells.join('')}</tr>`
    }

    const headingRow: string[] = []
    if (headings) {
      headingRow.push(
        `<tr style="height:${HEADING_ROW_HEIGHT_PT}pt"><th class="hd hd-top hd-left"></th>` +
          allColumns.map((c) => `<th class="hd hd-top">${columnLabel(c)}</th>`).join('') +
          `</tr>`,
      )
    }
    const head = headingRow.concat(
      titleRowList.map((row, index) => rowHtml(row, titleRowList, index)),
    )
    const body = pageRows.map((row, index) => rowHtml(row, pageRows, titleRowList.length + index))

    const originLeft = headingWidthPt + titleColumnsWidthPt
    const originTop = headingHeightPt + titleRowsHeightPt
    const bodyWidthPt = pageColumns.reduce((sum, c) => sum + columnWidthPt(c), 0)
    const bodyHeightPt = pageRows.reduce((sum, r) => sum + rowHeightPt(r), 0)
    const pageLeft = data.columnOffset.get(page.colStart) ?? 0
    const pageTop = data.rowOffset.get(page.rowStart) ?? 0
    const overlays = visuals.visuals
      .map((visual) => {
        const left = data.columnOffset.get(visual.fromColumn)
        const top = data.rowOffset.get(visual.fromRow)
        if (left === undefined || top === undefined) return ''
        const x = left - pageLeft + finitePt(visual.offsetXPx, 0) * 0.75
        const y = top - pageTop + finitePt(visual.offsetYPx, 0) * 0.75
        const widthPx = finitePt(visual.widthPx, 1)
        const heightPx = finitePt(visual.heightPx, 1)
        if (
          x >= bodyWidthPt ||
          y >= bodyHeightPt ||
          x + widthPx * 0.75 <= 0 ||
          y + heightPx * 0.75 <= 0
        )
          return ''
        return visualOverlayHtml(visual, x, y, widthPx, heightPx)
      })
      .join('')
    const overlayLayer = overlays
      ? `<div class="ov" style="left:${round(originLeft)}pt;top:${round(originTop)}pt;width:${round(bodyWidthPt)}pt;height:${round(bodyHeightPt)}pt">${overlays}</div>`
      : ''
    const table = `<table>${colgroup}<thead>${head.join('')}</thead><tbody>${body.join('')}</tbody></table>`
    return { body: table + overlayLayer, page }
  })

  return {
    sheetName,
    fileName,
    paper,
    margins,
    scale,
    headerFooterScale: setup.headerFooterScaleWithDoc ? scale : 1,
    pictures,
    headerFooter: {
      odd: { header: setup.header, footer: setup.footer },
      first: setup.firstPage,
      even: setup.evenPages,
    },
    css: visuals.css,
    pages: records,
    now: options.now ?? new Date(),
  }
}

export interface AssembledPrintJob {
  readonly html: string
  /// Index into `documents` for each emitted page.
  readonly pages: readonly { readonly document: number; readonly page: number }[]
}

/// The full print document: pages numbered across every sheet (&P/&N), each
/// page a paper-sized box with its header/footer; `include` keeps a subset
/// of the job's pages (page numbers still count the whole job, like Excel).
export function assemblePrintHtml(
  documents: readonly SheetPrintDocument[],
  include?: (jobPage: number) => boolean,
): AssembledPrintJob {
  const total = documents.reduce((sum, document) => sum + document.pages.length, 0)
  const paperClasses = new Map<string, string>()
  const pageRules: string[] = []
  const pages: string[] = []
  const emitted: { document: number; page: number }[] = []
  let jobPage = 0
  documents.forEach((document, documentIndex) => {
    const paperKey = `${document.paper.widthIn}x${document.paper.heightIn}`
    let paperClass = paperClasses.get(paperKey)
    if (paperClass === undefined) {
      paperClass = `p${paperClasses.size}`
      paperClasses.set(paperKey, paperClass)
      pageRules.push(
        `@page ${paperClass} { size: ${document.paper.widthIn}in ${document.paper.heightIn}in; margin: 0; }` +
          ` .${paperClass} { page: ${paperClass}; }`,
      )
    }
    document.pages.forEach((record, pageIndex) => {
      jobPage += 1
      if (include && !include(jobPage)) return
      emitted.push({ document: documentIndex, page: pageIndex })
      pages.push(
        pageHtml(
          document,
          record,
          jobPage,
          total,
          paperClass as string,
          pageVariant(document.headerFooter, pageIndex, jobPage),
        ),
      )
    })
  })
  const css = [...new Set(documents.map((document) => document.css).filter(Boolean))].join('\n')
  const lang = getLang()
  const html =
    `<!doctype html><html lang="${htmlLang(lang)}"><head><meta charset="utf-8"><style>
@page { margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: #fff; }
body { font-family: Calibri, 'Helvetica Neue', Arial, ${printCjkFonts(lang)}, sans-serif; }
.page { position: relative; overflow: hidden; background: #fff; break-after: page; }
.page:last-child { break-after: auto; }
.content { position: absolute; overflow: hidden; }
.hf { position: absolute; display: flex; color: #000; }
.hf span { flex: 1; min-width: 0; white-space: pre-wrap; }
.hf img { vertical-align: bottom; }
table { border-collapse: separate; border-spacing: 0; table-layout: fixed; }
td, th { padding: 0; border: 0; overflow: hidden; font-size: 11pt; vertical-align: bottom; }
td .c { display: flex; overflow: hidden; padding: 0 2pt; line-height: 1.2; }
td .c > span { width: 100%; }
th.hd { background: #f1f1f1; color: #444; font-size: 8.5pt; font-weight: 400;
  text-align: center; vertical-align: middle;
  box-shadow: inset -0.75pt 0 0 #b7b7b7, inset 0 -0.75pt 0 #b7b7b7; }
th.hd-top { box-shadow: inset -0.75pt 0 0 #b7b7b7, inset 0 -0.75pt 0 #b7b7b7, inset 0 0.75pt 0 #b7b7b7; }
th.hd-left { box-shadow: inset -0.75pt 0 0 #b7b7b7, inset 0 -0.75pt 0 #b7b7b7, inset 0.75pt 0 0 #b7b7b7; }
th.hd-top.hd-left { box-shadow: inset -0.75pt 0 0 #b7b7b7, inset 0 -0.75pt 0 #b7b7b7,
  inset 0 0.75pt 0 #b7b7b7, inset 0.75pt 0 0 #b7b7b7; }
.ov { position: absolute; overflow: hidden; }
.pv { position: absolute; overflow: hidden; }
.xlsx-print-visual { display: block; width: 100%; height: 100%; }
${pageRules.join('\n')}
</style>${css ? `<style>${css}</style>` : ''}</head><body>` +
    pages.join('') +
    `</body></html>`
  return { html, pages: emitted }
}

/// differentFirst is the sheet's first page; differentOddEven follows the
/// job-wide page number (the one &P prints).
export function pageVariant(
  headerFooter: SheetHeaderFooter,
  sheetPageIndex: number,
  jobPage: number,
): PageVariant {
  if (sheetPageIndex === 0 && headerFooter.first !== null) return 'first'
  if (jobPage % 2 === 0 && headerFooter.even !== null) return 'even'
  return 'odd'
}

function pageHtml(
  document: SheetPrintDocument,
  record: SheetPrintPageRecord,
  page: number,
  total: number,
  paperClass: string,
  variant: PageVariant,
): string {
  const { paper, margins } = document
  const pair =
    variant === 'first'
      ? (document.headerFooter.first ?? document.headerFooter.odd)
      : variant === 'even'
        ? (document.headerFooter.even ?? document.headerFooter.odd)
        : document.headerFooter.odd
  const fields = {
    page,
    total,
    date: document.now,
    fileName: document.fileName,
    sheetName: document.sheetName,
  }
  const header = pair.header
    ? headerFooterHtml(
        pair.header,
        'header',
        document,
        fields,
        sectionPictures(document.pictures, 'header', variant),
      )
    : ''
  const footer = pair.footer
    ? headerFooterHtml(
        pair.footer,
        'footer',
        document,
        fields,
        sectionPictures(document.pictures, 'footer', variant),
      )
    : ''
  const contentWidthIn = Math.max(paper.widthIn - margins.left - margins.right, 0.01)
  const contentHeightIn = Math.max(paper.heightIn - margins.top - margins.bottom, 0.01)
  return (
    `<div class="page ${paperClass}" style="width:${paper.widthIn}in;height:${paper.heightIn}in">` +
    header +
    `<div class="content" style="left:${round(margins.left)}in;top:${round(margins.top)}in;` +
    `width:${round(contentWidthIn)}in;height:${round(contentHeightIn)}in">` +
    `<div style="zoom:${round4(document.scale)};position:relative">${record.body}</div></div>` +
    footer +
    `</div>`
  )
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

const PICTURE_MARK = '￼'

/// One left/center/right header or footer inside the page: Excel offsets it
/// from the paper edge by the header/footer margin and spans the side
/// margins; the text and pictures follow the print scale (scaleWithDoc).
export function headerFooterHtml(
  parts: HeaderFooterParts,
  kind: 'header' | 'footer',
  document: Pick<SheetPrintDocument, 'margins' | 'headerFooterScale'>,
  fields: { page: number; total: number; date: Date; fileName: string; sheetName: string },
  pictures: SectionPictures = {},
): string {
  const sections = [parts.left ?? '', parts.center ?? '', parts.right ?? '']
  if (sections.every((text) => text === '')) return ''
  const sectionPicture = [pictures.left, pictures.center, pictures.right]
  const scale = document.headerFooterScale
  const rendered = sections.map((text, index) =>
    renderHeaderFooterSection(text, fields, sectionPicture[index], scale),
  )
  const { margins } = document
  const edge =
    kind === 'header' ? `top:${round(margins.header)}in` : `bottom:${round(margins.footer)}in`
  return (
    `<div class="hf" style="left:${round(margins.left)}in;right:${round(margins.right)}in;${edge};` +
    `font-size:${round(HEADER_FOOTER_FONT_SIZE_PT * scale)}pt">` +
    `<span>${rendered[0]}</span>` +
    `<span style="text-align:center">${rendered[1]}</span>` +
    `<span style="text-align:right">${rendered[2]}</span></div>`
  )
}

/// Field codes resolved for this page, escaped, with `&G` replaced by the
/// section's picture (nothing when the slot has none, like Excel).
export function renderHeaderFooterSection(
  text: string,
  fields: { page: number; total: number; date: Date; fileName: string; sheetName: string },
  picture?: HeaderFooterPictureImage,
  scale = 1,
): string {
  const expanded = expandHeaderFooterCodes(text, {
    ...fields,
    picture: picture ? PICTURE_MARK : '',
  })
  return expanded
    .split(PICTURE_MARK)
    .map(escapeHtml)
    .join(picture ? pictureHtml(picture, scale) : '')
}

/// The picture at its declared size times the print scale (points → CSS px
/// at 96/72). The data URL is built from a validated media type and base64
/// payload; escaping it anyway keeps the attribute closed no matter what.
function pictureHtml(picture: HeaderFooterPictureImage, scale: number): string {
  const width = round((picture.widthPt * scale * 96) / 72)
  const height = round((picture.heightPt * scale * 96) / 72)
  return `<img src="${escapeAttribute(picture.dataUrl)}" style="width:${width}px;height:${height}px">`
}

/// The export/print request for a job: the paper of the first sheet seeds
/// the dialog, the pages carry their own size via @page rules.
export function printRequest(
  documents: readonly SheetPrintDocument[],
  fileName: string,
  include?: (jobPage: number) => boolean,
): WorkbookExportPdfRequest {
  const job = assemblePrintHtml(documents, include)
  if (job.pages.length === 0) throw new PrintError(t('appPrintNothing'))
  const first = documents[job.pages[0]?.document ?? 0] ?? documents[0]
  const paper = first?.paper ?? paperGeometry(9, 'portrait')
  return {
    fileName,
    html: job.html,
    landscape: paper.landscape,
    pageSize: paper.pageSize,
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    scale: 1,
  }
}

/// Single-sheet export (headless CLI path and File → Export PDF).
export function buildSheetPrintPayload(
  worksheet: PrintWorksheet,
  setup: EffectivePageSetup,
  fileName: string,
  sheetName: string,
  pictures: HeaderFooterPictures = new Map(),
  visuals: PrintVisualSnapshot = { visuals: [], css: '' },
  options: SheetPrintOptions = {},
): WorkbookExportPdfRequest {
  const baseName = fileName.replace(/\.pdf$/, '')
  return printRequest(
    [layoutSheetPrint(worksheet, setup, baseName, sheetName, pictures, visuals, options)],
    fileName,
  )
}

/// Excel's default print range covers the cells and the drawings over them.
function usedArea(worksheet: PrintWorksheet, visuals: readonly PrintVisual[]): PrintArea {
  return {
    startRow: 0,
    startColumn: 0,
    endRow: Math.max(worksheet.getLastRow(), 0, ...visuals.map((visual) => visual.toRow)),
    endColumn: Math.max(worksheet.getLastColumn(), 0, ...visuals.map((visual) => visual.toColumn)),
  }
}

/// The snapshot at its anchor. Columns are laid out at px × 0.75 pt, which
/// is one CSS px per sheet px, so the clone keeps its px box inside a box of
/// the same size stated in pt.
function visualOverlayHtml(
  visual: PrintVisual,
  leftPt: number,
  topPt: number,
  widthPx: number,
  heightPx: number,
): string {
  const style = `left:${round(leftPt)}pt;top:${round(topPt)}pt;width:${round(widthPx * 0.75)}pt;height:${round(heightPx * 0.75)}pt`
  const inner = `width:${round(widthPx)}px;height:${round(heightPx)}px`
  return `<div class="pv" style="${style}"><div style="${inner}">${visual.html}</div></div>`
}

function parseArea(area: string): PrintArea {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(area)
  if (!match) throw new PrintError(t('appPrintBadArea', { area }))
  const rows = [Number(match[2]) - 1, Number(match[4]) - 1]
  const columns = [columnIndex(match[1] ?? 'A'), columnIndex(match[3] ?? 'A')]
  return {
    startRow: Math.min(...rows),
    endRow: Math.max(...rows),
    startColumn: Math.min(...columns),
    endColumn: Math.max(...columns),
  }
}

function parseTitleRows(titles: string): IndexSpan {
  const match = /^(\d{1,7}):(\d{1,7})$/.exec(titles)
  if (!match) throw new PrintError(t('appPrintBadTitles', { titles }))
  const start = Number(match[1]) - 1
  const end = Number(match[2]) - 1
  if (end - start > 20) throw new PrintError(t('appPrintTitlesLimit'))
  return { start, end }
}

function parseTitleColumns(titles: string): IndexSpan {
  const match = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(titles)
  if (!match) throw new PrintError(t('appPrintBadTitles', { titles }))
  const start = columnIndex((match[1] ?? 'A').toUpperCase())
  const end = columnIndex((match[2] ?? 'A').toUpperCase())
  if (end - start > 20) throw new PrintError(t('appPrintTitlesLimit'))
  return { start, end }
}

function spanIndices(span: IndexSpan): number[] {
  return Array.from({ length: span.end - span.start + 1 }, (_, offset) => span.start + offset)
}

function spanTotal(span: IndexSpan | null, sizeOf: (index: number) => number): number {
  return span ? spanIndices(span).reduce((sum, index) => sum + sizeOf(index), 0) : 0
}

/// Running offset of each body row/column from the area's first one.
function offsets(
  indices: readonly number[],
  sizeOf: (index: number) => number,
): Map<number, number> {
  const result = new Map<number, number>()
  let position = 0
  for (const index of indices) {
    result.set(index, position)
    position += sizeOf(index)
  }
  return result
}

function mergeMaps(worksheet: PrintWorksheet, area: PrintArea) {
  const anchors = new Map<string, { rows: number; columns: number }>()
  const covered = new Map<string, readonly [number, number]>()
  for (const merge of worksheet.getMergedRanges()) {
    const row = merge.getRow()
    const column = merge.getColumn()
    if (row > area.endRow || column > area.endColumn) continue
    if (row + merge.getHeight() - 1 < area.startRow) continue
    if (column + merge.getWidth() - 1 < area.startColumn) continue
    anchors.set(`${row}:${column}`, { rows: merge.getHeight(), columns: merge.getWidth() })
    for (let r = row; r < row + merge.getHeight(); r += 1) {
      for (let c = column; c < column + merge.getWidth(); c += 1) {
        if (r !== row || c !== column) covered.set(`${r}:${c}`, [row, column])
      }
    }
  }
  return { anchors, covered }
}

function cellDisplay(worksheet: PrintWorksheet, row: number, column: number): string {
  return worksheet.getRange(row, column, 1, 1).getDisplayValues()[0]?.[0] ?? ''
}

/// Borders paint as inset shadows (no layout share, unlike CSS borders, so
/// the planned row heights hold); `edges` says which of the cell's four
/// edges it owns.
function cellCss(
  style: PrintCellStyle | null,
  rawValue: unknown,
  gridlines: boolean,
  edges: { top: boolean; bottom: boolean; left: boolean; right: boolean },
): string {
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
  rules.push(style?.tb === 3 ? 'white-space:pre-wrap;word-break:break-word' : 'white-space:pre')
  const shadows: string[] = []
  const edge = (
    key: 't' | 'b' | 'l' | 'r',
    draw: boolean,
    offset: (width: number) => string,
  ): void => {
    const border = style?.bd?.[key]
    if (border && draw) {
      shadows.push(
        `inset ${offset(printBorderWidthPt(border.s))} 0 ${cssColor(border.cl?.rgb ?? '#000000')}`,
      )
    } else if (gridlines && draw) {
      shadows.push(`inset ${offset(GRIDLINE_PT)} 0 #c0c0c0`)
    }
  }
  edge('t', edges.top, (width) => `0 ${width}pt`)
  edge('b', edges.bottom, (width) => `0 -${width}pt`)
  edge('l', edges.left, (width) => `${width}pt 0`)
  edge('r', edges.right, (width) => `-${width}pt 0`)
  if (shadows.length > 0) rules.push(`box-shadow:${shadows.join(',')}`)
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

function round4(value: number): number {
  return Math.round(value * 10000) / 10000
}
