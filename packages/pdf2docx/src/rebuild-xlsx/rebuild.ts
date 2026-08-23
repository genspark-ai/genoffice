/**
 * IR → xlsx workbook assembly (P26). One PDF page = one worksheet ("Page N",
 * Adobe's convention). Tables are the star: every TableBlock lands as a cell
 * grid in reading order (gridSpan/vMerge → mergeCells, fills/borders/vAlign
 * preserved, numeric-looking text stored as real numbers). Non-table text is
 * kept — one paragraph per row in column A — so no content is lost. Images
 * and decor stay out (v1: no floating pictures in spreadsheets).
 */
import { HF_PAGE_MARK, type FurnitureHf } from '../analyze/furniture'
import type { Span, TableBlock, TableCellBlock, TextBlock, IrPage } from '../ir'
import { blockRuns } from '../rebuild-pptx/text'
import { parseCellValue } from './numbers'
import { splitBandRows } from './rowsplit'
import {
  buildXlsxPackage,
  cellRef,
  StylePool,
  type AlignmentSpec,
  type BorderEdges,
  type CellStyleSpec,
  type FontSpec,
  type SheetCell,
  type SheetSpec,
} from './workbook'

/**
 * pt → Excel column width (character units of the default font): width
 * pixels = chars × 7 + 5 at 96 dpi, and 1 pt = 4/3 px, so
 * chars = (pt × 4/3 − 5) / 7. Verified against LO Calc renders (a 96 pt
 * PDF column round-trips to ~23 chars ≈ 96 pt on screen).
 */
export function ptToColumnChars(pt: number): number {
  return Math.max(2, (pt * (4 / 3) - 5) / 7)
}

/** xlsx hard limit for row height */
const MAX_ROW_HEIGHT_PT = 409

// ── text flattening ──

/** flatten one text block to plain cell text (pptx line-join semantics) */
export function flattenBlockText(block: TextBlock): string {
  let text = blockRuns(block, 1)
    .map((r) => r.text)
    .join('')
  if (block.list?.marker !== undefined && text !== '') {
    text = `${block.list.marker.trimEnd()} ${text}`
  }
  if (block.tocEntry && text !== '') text = `${text}\t${block.tocEntry.pageNumber}`
  return text
}

/** dominant run style by character count (cell-level rich text is out of v1) */
function dominantFont(blocks: TextBlock[]): FontSpec | undefined {
  const weights = new Map<string, { font: FontSpec; chars: number }>()
  for (const block of blocks) {
    for (const line of block.lines) {
      for (const span of line.spans) {
        if (span.invisible || span.text === '') continue
        const font = fontOfSpan(span)
        const key = JSON.stringify(font)
        const entry = weights.get(key)
        if (entry) entry.chars += span.text.length
        else weights.set(key, { font, chars: span.text.length })
      }
    }
  }
  let best: { font: FontSpec; chars: number } | undefined
  for (const entry of weights.values()) {
    if (!best || entry.chars > best.chars) best = entry
  }
  return best?.font
}

function fontOfSpan(span: Span): FontSpec {
  const font: FontSpec = {}
  if (span.bold) font.bold = true
  if (span.italic) font.italic = true
  if (span.color && span.color !== '000000') font.color = span.color
  if (span.fontSize > 0) font.sizePt = Math.round(span.fontSize * 2) / 2
  if (span.fontFamily) font.name = span.fontFamily
  return font
}

// ── table grid helpers (same derivations as rebuild-pptx/table.ts) ──

/** column start index of every cell in one IR row (gridSpan advances) */
function colStarts(row: TableCellBlock[]): number[] {
  const starts: number[] = []
  let col = 0
  for (const cell of row) {
    starts.push(col)
    col += Math.max(1, cell.gridSpan)
  }
  return starts
}

/** rows a vMerge-restart cell spans: 1 + following continue placeholders in its column */
function rowSpanOf(block: TableBlock, rowIdx: number, colStart: number): number {
  let span = 1
  for (let r = rowIdx + 1; r < block.rows.length; r++) {
    const row = block.rows[r]!
    const starts = colStarts(row)
    const idx = starts.indexOf(colStart)
    if (idx < 0 || row[idx]!.vMerge !== 'continue') break
    span++
  }
  return span
}

/** per-row heights from row-top boundaries (first non-continue cell's box top) */
function tableRowHeightsPt(block: TableBlock): number[] {
  const tops: number[] = []
  for (const [r, row] of block.rows.entries()) {
    const anchor = row.find((c) => c.vMerge !== 'continue') ?? row[0]
    tops.push(anchor ? anchor.box.y1 : block.box.y1 - r)
  }
  const heights: number[] = []
  for (let r = 0; r < tops.length; r++) {
    const bottom = r + 1 < tops.length ? tops[r + 1]! : block.box.y0
    heights.push(Math.max(1, tops[r]! - bottom))
  }
  return heights
}

const totalColumns = (block: TableBlock): number => block.colWidthsPt.length

/** cumulative x boundaries of a table's grid columns */
function tableBounds(block: TableBlock): number[] {
  const bounds: number[] = [block.box.x0]
  for (const w of block.colWidthsPt) bounds.push(bounds[bounds.length - 1]! + w)
  return bounds
}

// ── cross-page table runs (P39) ──

/** column boundaries and table x-range must agree within this to continue a run */
const RUN_GEOM_TOL_PT = 2
/** a run header needs at least this many non-empty cells (single-cell boxes never chain) */
const RUN_HEADER_MIN_CELLS = 2

const tableCellText = (cell: TableCellBlock): string => cell.blocks.map(flattenBlockText).join('\n')

const headerTexts = (t: TableBlock): string[] =>
  (t.rows[0] ?? []).map((c) => tableCellText(c).replace(/\s+/g, ' ').trim())

const firstTableOf = (page: IrPage): TableBlock | undefined =>
  page.blocks.find((b): b is TableBlock => b.kind === 'table')

const lastTableOf = (page: IrPage): TableBlock | undefined => {
  for (let i = page.blocks.length - 1; i >= 0; i--) {
    const b = page.blocks[i]!
    if (b.kind === 'table') return b
  }
  return undefined
}

/**
 * Statement-style continuation: the next page restates the SAME table —
 * identical column boundaries and an identical repeated header row (bank
 * statements, transaction logs). Geometry alone is not enough: unrelated
 * same-width grids on consecutive pages must not chain.
 */
function tableContinues(a: TableBlock, b: TableBlock): boolean {
  if (a.colWidthsPt.length !== b.colWidthsPt.length) return false
  if (Math.abs(a.box.x0 - b.box.x0) > RUN_GEOM_TOL_PT) return false
  if (Math.abs(a.box.x1 - b.box.x1) > RUN_GEOM_TOL_PT) return false
  for (let i = 0; i < a.colWidthsPt.length; i++) {
    if (Math.abs(a.colWidthsPt[i]! - b.colWidthsPt[i]!) > RUN_GEOM_TOL_PT) return false
  }
  const ha = headerTexts(a)
  const hb = headerTexts(b)
  if (ha.filter((t) => t !== '').length < RUN_HEADER_MIN_CELLS) return false
  return ha.length === hb.length && ha.every((t, i) => t === hb[i])
}

/** consecutive pages whose last/first tables continue merge into one worksheet */
export function planPageGroups(pages: IrPage[]): number[][] {
  const groups: number[][] = []
  for (let i = 0; i < pages.length; i++) {
    const page = pages[i]!
    const prev = i > 0 ? pages[i - 1] : undefined
    const group = groups[groups.length - 1]
    if (group && prev && !prev.scanned && !prev.degraded && !page.scanned && !page.degraded) {
      const a = lastTableOf(prev)
      const b = firstTableOf(page)
      if (a && b && tableContinues(a, b)) {
        group.push(i)
        continue
      }
    }
    groups.push([i])
  }
  return groups
}

// ── per-page emission ──

interface SheetBuilder {
  cells: SheetCell[]
  merges: string[]
  rowHeightsPt: Map<number, number>
  /** per-column width claims in pt (max across the page's tables) */
  colWidthsPt: number[]
  row: number
  /** page x-position → sheet column mapping (P27) */
  slots: number[]
}

/** x-anchors closer than this merge into one sheet column (chained clustering;
 * a lattice grid's own boundaries are ≥MIN_CELL_DIM=3pt apart and never chain) */
const SLOT_TOL_PT = 1.5

/**
 * Plan the sheet's columns from page geometry (P27): every text block's left
 * edge and every table's column boundaries become x-slots; each slot is one
 * sheet column. A centered 2-column table beside margin prose then lands in
 * its own columns instead of overwriting column A — the sheet mirrors the
 * page like Adobe's PDF→Excel export.
 */
function planColumnSlots(
  pages: readonly IrPage[],
  boundsFor?: Map<TableBlock, number[]>,
): number[] {
  const xs: number[] = []
  for (const page of pages) {
    for (const block of page.blocks) {
      if (block.kind === 'image') continue
      if (block.kind === 'table') {
        xs.push(...(boundsFor?.get(block) ?? tableBounds(block)))
      } else {
        xs.push(block.box.x0)
      }
    }
    for (const note of page.footnotes ?? []) {
      for (const block of note.blocks) xs.push(block.box.x0)
    }
  }
  xs.sort((a, b) => a - b)
  const slots: number[] = []
  for (const x of xs) {
    const last = slots[slots.length - 1]
    if (last === undefined || x - last > SLOT_TOL_PT) slots.push(x)
  }
  return slots
}

/** sheet column of a page x position (nearest slot; slots cover every anchor) */
function slotOf(slots: number[], x: number): number {
  let best = 0
  let bestDist = Infinity
  for (const [i, s] of slots.entries()) {
    const d = Math.abs(s - x)
    if (d < bestDist) {
      bestDist = d
      best = i
    }
  }
  return best
}

/**
 * borders follow the docx/pptx rebuild semantics: lattice tables (drawn
 * rulings, confidence absent) get the full grid; rule-separated zones keep
 * only their inside-vertical separator; stream/form tables stay borderless.
 */
function cellBorder(
  block: TableBlock,
  colStart: number,
  span: number,
  softEdges?: { left?: boolean; right?: boolean; top?: boolean; bottom?: boolean },
): BorderEdges | undefined {
  if (block.confidence === undefined) {
    const color = block.borderColor ?? '000000'
    const edges: BorderEdges = { top: color, right: color, bottom: color, left: color }
    // split-run cells (P27) / split band rows (P40): the rule was never drawn
    if (softEdges?.left) delete edges.left
    if (softEdges?.right) delete edges.right
    if (softEdges?.top) delete edges.top
    if (softEdges?.bottom) delete edges.bottom
    return edges
  }
  if (block.sepRule) {
    const color = block.borderColor ?? '000000'
    const edges: BorderEdges = {}
    if (colStart > 0) edges.left = color
    if (colStart + span < totalColumns(block)) edges.right = color
    return edges.left || edges.right ? edges : undefined
  }
  return undefined
}

function emitTable(
  sheet: SheetBuilder,
  styles: StylePool,
  block: TableBlock,
  dropHeaderRow = false,
  boundsOverride?: number[],
): void {
  const startRow = sheet.row
  const heights = tableRowHeightsPt(block)
  // grid column k lives at page x boundary[k] → sheet column slot (P27);
  // continuation tables map through the run anchor's boundaries so pairwise
  // drift within RUN_GEOM_TOL_PT can never land rows one slot off (P39)
  const bounds = boundsOverride ?? tableBounds(block)
  const sheetColOf = (gridCol: number): number => slotOf(sheet.slots, bounds[gridCol]!)
  for (const [i, w] of block.colWidthsPt.entries()) {
    const sc = sheetColOf(i)
    sheet.colWidthsPt[sc] = Math.max(sheet.colWidthsPt[sc] ?? 0, w)
  }

  const skip = dropHeaderRow ? 1 : 0
  for (const [r, row] of block.rows.entries()) {
    if (r < skip) continue
    const sheetRow = startRow + r - skip
    const height = heights[r]
    if (height !== undefined) {
      sheet.rowHeightsPt.set(sheetRow, Math.min(MAX_ROW_HEIGHT_PT, height))
    }
    const starts = colStarts(row)
    for (const [i, cell] of row.entries()) {
      const colStart = starts[i]!
      const span = Math.max(1, cell.gridSpan)
      const border = cellBorder(block, colStart, span, cell.softEdges)
      const sheetCol = sheetColOf(colStart)
      // the spanned grid columns cover sheet slots [sheetCol, sheetColEnd]
      const sheetColEnd = Math.max(sheetCol, sheetColOf(colStart + span) - 1)
      const sheetSpan = sheetColEnd - sheetCol + 1

      if (cell.vMerge === 'continue') {
        // covered placeholder: no content, but the style must exist so the
        // grid's borders stay closed on every covered position
        emitCoveredCells(sheet, styles, sheetRow, sheetCol, sheetSpan, cell.fill, border)
        continue
      }

      const rowSpan = cell.vMerge === 'restart' ? rowSpanOf(block, r, colStart) : 1
      if (sheetSpan > 1 || rowSpan > 1) {
        sheet.merges.push(
          `${cellRef(sheetRow, sheetCol)}:${cellRef(sheetRow + rowSpan - 1, sheetColEnd)}`,
        )
      }

      const text = cell.blocks.map(flattenBlockText).join('\n')
      const parsed = parseCellValue(text)
      const align: AlignmentSpec = { vertical: cell.vAlign ?? 'top' }
      if (text.includes('\n')) align.wrapText = true
      if (parsed.kind === 'text') {
        const blockAlign = cell.blocks[0]?.align
        if (blockAlign === 'center' || blockAlign === 'right') align.horizontal = blockAlign
      }
      const style: CellStyleSpec = { align }
      const font = dominantFont(cell.blocks)
      if (font && Object.keys(font).length > 0) style.font = font
      if (cell.fill) style.fill = cell.fill
      if (border) style.border = border
      if (parsed.kind === 'number' && parsed.numFmt !== 'General') style.numFmt = parsed.numFmt

      const spec: SheetCell = { row: sheetRow, col: sheetCol, styleId: styles.cellXf(style) }
      if (text !== '') {
        spec.value =
          parsed.kind === 'number'
            ? { kind: 'number', value: parsed.value }
            : { kind: 'text', text }
      }
      sheet.cells.push(spec)
      // trailing spanned columns of the anchor row are covered placeholders too
      emitCoveredCells(sheet, styles, sheetRow, sheetCol + 1, sheetSpan - 1, cell.fill, border)
    }
  }
  sheet.row = startRow + block.rows.length - skip
}

function emitCoveredCells(
  sheet: SheetBuilder,
  styles: StylePool,
  row: number,
  colStart: number,
  count: number,
  fill: string | undefined,
  border: BorderEdges | undefined,
): void {
  for (let c = 0; c < count; c++) {
    const style: CellStyleSpec = {}
    if (fill) style.fill = fill
    if (border) style.border = border
    sheet.cells.push({ row, col: colStart + c, styleId: styles.cellXf(style) })
  }
}

function emitTextRow(sheet: SheetBuilder, styles: StylePool, block: TextBlock): void {
  const text = flattenBlockText(block)
  if (text.trim() === '') return
  const parsed = parseCellValue(text)
  const style: CellStyleSpec = {}
  const font = dominantFont([block])
  if (font && Object.keys(font).length > 0) style.font = font
  if (text.includes('\n')) style.align = { wrapText: true }
  if (parsed.kind === 'number' && parsed.numFmt !== 'General') style.numFmt = parsed.numFmt
  const col = slotOf(sheet.slots, block.box.x0)
  const spec: SheetCell = { row: sheet.row, col, styleId: styles.cellXf(style) }
  spec.value =
    parsed.kind === 'number' ? { kind: 'number', value: parsed.value } : { kind: 'text', text }
  sheet.cells.push(spec)
  sheet.row += 1
}

function emitGroup(groupPages: readonly IrPage[], styles: StylePool): SheetSpec {
  const first = groupPages[0]!
  const last = groupPages[groupPages.length - 1]!
  const name =
    groupPages.length === 1
      ? `Page ${first.index + 1}`
      : `Pages ${first.index + 1}-${last.index + 1}`
  // continuation tables borrow their run anchor's column boundaries: pairwise
  // drift within RUN_GEOM_TOL_PT would otherwise mint near-duplicate slots
  // (SLOT_TOL_PT is tighter) and shift continuation rows one column over.
  // Anchors propagate along the chain — a run can start mid-group (the page
  // continuing table A may end with a NEW table B that the next page continues)
  const runBounds = new Map<TableBlock, number[]>()
  for (let pi = 1; pi < groupPages.length; pi++) {
    const prevLast = lastTableOf(groupPages[pi - 1]!)
    const t = firstTableOf(groupPages[pi]!)
    if (prevLast && t) runBounds.set(t, runBounds.get(prevLast) ?? tableBounds(prevLast))
  }
  const sheet: SheetBuilder = {
    cells: [],
    merges: [],
    rowHeightsPt: new Map(),
    colWidthsPt: [],
    row: 0,
    slots: planColumnSlots(groupPages, runBounds),
  }

  // scanned/degraded pages never continue a run — always a singleton group
  if (first.scanned || first.degraded) {
    const kind = first.scanned
      ? 'scanned page'
      : `degraded page (${first.degradedReason ?? 'unknown'})`
    sheet.cells.push({
      row: 0,
      col: 0,
      styleId: 0,
      value: {
        kind: 'text',
        text: `Page ${first.index + 1}: ${kind} — not convertible to cells`,
      },
    })
    return { name, cells: sheet.cells }
  }

  let lastWasTable = false
  for (const [pi, page] of groupPages.entries()) {
    // the restated header row of the continuation table is dropped so the
    // run reads as one table; its rows attach without a spacer
    const continuation = pi > 0 ? firstTableOf(page) : undefined
    for (const block of page.blocks) {
      if (block.kind === 'image') continue
      if (block.kind === 'table') {
        if (block === continuation) {
          emitTable(sheet, styles, splitBandRows(block), true, runBounds.get(block))
        } else {
          // one blank spacer row between adjacent tables
          if (lastWasTable) sheet.row += 1
          emitTable(sheet, styles, splitBandRows(block))
        }
        lastWasTable = true
      } else {
        emitTextRow(sheet, styles, block)
        lastWasTable = false
      }
    }
    for (const note of page.footnotes ?? []) {
      for (const block of note.blocks) emitTextRow(sheet, styles, block)
    }
  }

  const spec: SheetSpec = { name, cells: sheet.cells }
  if (sheet.merges.length > 0) spec.merges = sheet.merges
  if (sheet.rowHeightsPt.size > 0) spec.rowHeightsPt = sheet.rowHeightsPt
  if (sheet.colWidthsPt.length > 0) {
    spec.colWidths = sheet.colWidthsPt.map((w) => (w > 0 ? ptToColumnChars(w) : undefined))
  }
  return spec
}

export interface RebuildXlsxResult {
  xlsx: Uint8Array
  /** notes appended by the xlsx layer itself (e.g. no tables in the document) */
  warnings: string[]
  /** the serialized sheet model (evaluation hook: cell-level ground truth) */
  sheets: SheetSpec[]
}

/** Rebuild the analyzed pages as an xlsx workbook (one worksheet per page). */
/** portrait [width, height] (pt) → OOXML paperSize code */
const PAPER_SIZES: Array<[number, number, number]> = [
  [612, 792, 1], // Letter
  [612, 1008, 5], // Legal
  [842, 1191, 8], // A3
  [595, 842, 9], // A4
  [420, 595, 11], // A5
  [516, 729, 13], // B5 (JIS)
]
const PAPER_TOL_PT = 3

/** source-page print geometry; paperSize omitted for non-standard pages */
function paperOf(page: IrPage | undefined): {
  paperSize?: number
  orientation: 'portrait' | 'landscape'
} {
  if (!page) return { orientation: 'portrait' }
  const w = Math.min(page.widthPt, page.heightPt)
  const h = Math.max(page.widthPt, page.heightPt)
  const match = PAPER_SIZES.find(
    ([pw, ph]) => Math.abs(pw - w) <= PAPER_TOL_PT && Math.abs(ph - h) <= PAPER_TOL_PT,
  )
  return {
    ...(match ? { paperSize: match[2] } : {}),
    orientation: page.widthPt > page.heightPt ? 'landscape' : 'portrait',
  }
}

/** one furniture line, &-encoded for an xlsx header/footer section */
const hfLine = (hf: FurnitureHf): string =>
  hf.text.replaceAll('&', '&&').replaceAll(HF_PAGE_MARK, '&P')

/**
 * Repeated PDF furniture (running headers/footers, page numbers, per-page
 * disclaimers) is dropped from the body text — worksheets carry it as the
 * PRINT header/footer instead, positioned &L/&C/&R by its ink span.
 */
function headerFooterOf(
  furnitureHf: readonly FurnitureHf[],
  pageWidthPt: number,
): SheetSpec['headerFooter'] {
  const bands = {
    top: { L: [] as string[], C: [] as string[], R: [] as string[] },
    bottom: { L: [] as string[], C: [] as string[], R: [] as string[] },
  }
  // reading order within a band: top lines by distance from the top edge,
  // bottom lines by descending distance from the bottom edge
  const ordered = [...furnitureHf].sort((a, b) =>
    a.band === 'top' ? a.edgeDistPt - b.edgeDistPt : b.edgeDistPt - a.edgeDistPt,
  )
  for (const hf of ordered) {
    const cx = (hf.x0 + hf.x1) / 2
    const sec = cx < pageWidthPt / 3 ? 'L' : cx > (2 * pageWidthPt) / 3 ? 'R' : 'C'
    bands[hf.band][sec].push(hfLine(hf))
  }
  const encode = (b: { L: string[]; C: string[]; R: string[] }): string =>
    (b.L.length ? `&L${b.L.join('\n')}` : '') +
    (b.C.length ? `&C${b.C.join('\n')}` : '') +
    (b.R.length ? `&R${b.R.join('\n')}` : '')
  const oddHeader = encode(bands.top)
  const oddFooter = encode(bands.bottom)
  if (!oddHeader && !oddFooter) return undefined
  return {
    ...(oddHeader ? { oddHeader } : {}),
    ...(oddFooter ? { oddFooter } : {}),
  }
}

export async function rebuildXlsx(
  pages: IrPage[],
  furnitureHf: readonly FurnitureHf[] = [],
): Promise<RebuildXlsxResult> {
  const styles = new StylePool()
  const groups = planPageGroups(pages)
  const sheets: SheetSpec[] = groups.map((g) =>
    emitGroup(
      g.map((i) => pages[i]!),
      styles,
    ),
  )
  // an empty workbook is invalid — a PDF with zero pages still ships one sheet
  if (sheets.length === 0) sheets.push({ name: 'Page 1', cells: [] })
  const pageWidthPt = pages[0]?.widthPt ?? 612
  const headerFooter = headerFooterOf(furnitureHf, pageWidthPt)
  // slots that skip the document's first page (cover sheets) stay off sheet 1
  const firstPageHf = headerFooterOf(
    furnitureHf.filter((hf) => hf.coversFirstPage),
    pageWidthPt,
  )
  if (headerFooter) {
    const trackedPageNo = furnitureHf.some((hf) => hf.pageNo)
    for (const [gi, sheet] of sheets.entries()) {
      const firstPage = groups[gi]?.[0] ?? 0
      const hf = firstPage === 0 ? firstPageHf : headerFooter
      if (!hf) continue
      sheet.headerFooter = hf
      // writing pageSetup at all defaults the omitted paper geometry to
      // Letter/portrait — carry the source page's instead; &P restarts at 1
      // on every worksheet, so pin the printed number to the group's first page
      sheet.pageSetup = {
        ...paperOf(pages[firstPage]),
        ...(trackedPageNo ? { firstPageNumber: firstPage + 1 } : {}),
      }
    }
  }

  const warnings: string[] = []
  const hasTable = pages.some((p) => p.blocks.some((b) => b.kind === 'table'))
  const hasOkPage = pages.some((p) => !p.scanned && !p.degraded)
  if (!hasTable && hasOkPage) warnings.push('no tables detected')

  const xlsx = await buildXlsxPackage(sheets, styles)
  return { xlsx, warnings, sheets }
}
