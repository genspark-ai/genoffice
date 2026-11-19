/// Formula auditing overlays: Trace Precedents / Dependents. References are
/// parsed out of formula text (same tokenizer as the streaming closure) and
/// drawn as arrows in float DOM layers anchored to cell ranges,
/// so they scroll and zoom with the grid.

import type { createUniver } from './create-univer'

import { qualifierMatches } from '@genoffice/xlsx-gateway/gateway/xlsx-structure'
import { parseFormulaReferences } from './formula-closure'

type UniverRuntime = ReturnType<typeof createUniver>
type ActiveWorkbook = NonNullable<ReturnType<UniverRuntime['univerAPI']['getActiveWorkbook']>>
type UniverWorksheet = NonNullable<ReturnType<ActiveWorkbook['getActiveSheet']>>

interface Disposable {
  dispose(): void
}

export interface CellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export interface CellPoint {
  readonly row: number
  readonly column: number
}

export interface SheetBounds {
  readonly lastRow: number
  readonly lastColumn: number
  readonly maxRows: number
  readonly maxColumns: number
}

export interface PrecedentResult {
  /// Same-sheet referenced ranges, deduped, in formula order.
  readonly areas: CellArea[]
  /// References that resolve to another sheet — counted, not drawn.
  readonly offSheet: number
}

export function collectPrecedents(
  formula: string,
  sheetName: string,
  bounds: SheetBounds,
): PrecedentResult {
  const areas: CellArea[] = []
  const seen = new Set<string>()
  let offSheet = 0
  for (const reference of parseFormulaReferences(formula)) {
    if (reference.qualifier !== undefined && !qualifierMatches(reference.qualifier, sheetName)) {
      offSheet += 1
      continue
    }
    // Whole row/column references clamp to the used range so the overlay
    // stays finite; explicit refs clamp only to the sheet extent.
    const area = {
      startRow: Math.max(reference.token.startRow ?? 0, 0),
      endRow: Math.min(reference.token.endRow ?? Math.max(bounds.lastRow, 0), bounds.maxRows - 1),
      startColumn: Math.max(reference.token.startColumn ?? 0, 0),
      endColumn: Math.min(
        reference.token.endColumn ?? Math.max(bounds.lastColumn, 0),
        bounds.maxColumns - 1,
      ),
    }
    if (area.endRow < area.startRow || area.endColumn < area.startColumn) continue
    const key = `${area.startRow}:${area.startColumn}:${area.endRow}:${area.endColumn}`
    if (seen.has(key)) continue
    seen.add(key)
    areas.push(area)
  }
  return { areas, offSheet }
}

type CellDataMatrix = Record<
  number,
  | Record<
      number,
      | {
          readonly f?: unknown
        }
      | null
      | undefined
    >
  | null
  | undefined
>

export interface DependentsSource {
  readonly sheets: Record<
    string,
    | {
        readonly name?: string | undefined
        readonly cellData?: CellDataMatrix | undefined
      }
    | undefined
  >
}

export interface DependentResult {
  /// Formula cells on the active sheet that read the target cell.
  readonly cells: CellPoint[]
  /// Dependent formulas living on other sheets — counted, not drawn.
  readonly offSheet: number
}

export function collectDependents(
  source: DependentsSource,
  activeSheetName: string,
  target: CellPoint,
): DependentResult {
  const cells: CellPoint[] = []
  let offSheet = 0
  for (const sheet of Object.values(source.sheets)) {
    if (!sheet?.cellData) continue
    const onActiveSheet = sheet.name === activeSheetName
    for (const [rowKey, rowCells] of Object.entries(sheet.cellData)) {
      for (const [columnKey, cell] of Object.entries(rowCells ?? {})) {
        const formula = cell?.f
        if (typeof formula !== 'string' || !formula.startsWith('=')) continue
        const row = Number(rowKey)
        const column = Number(columnKey)
        if (onActiveSheet && row === target.row && column === target.column) continue
        if (!formulaReadsCell(formula, onActiveSheet, activeSheetName, target)) continue
        if (onActiveSheet) cells.push({ row, column })
        else offSheet += 1
      }
    }
  }
  return { cells, offSheet }
}

function formulaReadsCell(
  formula: string,
  onActiveSheet: boolean,
  activeSheetName: string,
  target: CellPoint,
): boolean {
  for (const reference of parseFormulaReferences(formula)) {
    const targetsActiveSheet =
      reference.qualifier === undefined
        ? onActiveSheet
        : qualifierMatches(reference.qualifier, activeSheetName)
    if (!targetsActiveSheet) continue
    const { token } = reference
    if (
      (token.startRow ?? 0) <= target.row &&
      target.row <= (token.endRow ?? Number.MAX_SAFE_INTEGER) &&
      (token.startColumn ?? 0) <= target.column &&
      target.column <= (token.endColumn ?? Number.MAX_SAFE_INTEGER)
    ) {
      return true
    }
  }
  return false
}

export interface TraceArrow {
  readonly from: CellPoint
  readonly to: CellPoint
}

const ARROW_COLOR = '#1f5aa8'

/// Installs one float layer per arrow (spanning the from→to bounding box)
/// plus a border box per highlighted range. Pixel geometry is sampled from
/// current row heights / column widths; zoom scales the anchored layer and
/// the SVG viewBox together, so coordinates stay valid.
export function installTraceArrows(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  arrows: readonly TraceArrow[],
  highlights: readonly CellArea[],
  idPrefix: string,
): Disposable[] {
  const disposables: Disposable[] = []
  highlights.forEach((area, index) => {
    const key = `${idPrefix}-box-${index}`
    disposables.push(
      runtime.univerAPI.registerComponent(key, () => <div className="trace-precedent-box" />),
    )
    const floating = worksheet.addFloatDomToRange(
      worksheet.getRange(
        area.startRow,
        area.startColumn,
        area.endRow - area.startRow + 1,
        area.endColumn - area.startColumn + 1,
      ),
      { componentKey: key, allowTransform: false, eventPassThrough: true },
      {},
      key,
    )
    if (floating) disposables.push(floating)
  })
  arrows.forEach((arrow, index) => {
    if (arrow.from.row === arrow.to.row && arrow.from.column === arrow.to.column) return
    const box = {
      startRow: Math.min(arrow.from.row, arrow.to.row),
      endRow: Math.max(arrow.from.row, arrow.to.row),
      startColumn: Math.min(arrow.from.column, arrow.to.column),
      endColumn: Math.max(arrow.from.column, arrow.to.column),
    }
    const centerX = axisCenters(box.startColumn, box.endColumn, (i) => worksheet.getColumnWidth(i))
    const centerY = axisCenters(box.startRow, box.endRow, (i) => worksheet.getRowHeight(i))
    const geometry = {
      width: centerX.total,
      height: centerY.total,
      x1: centerX.at(arrow.from.column),
      y1: centerY.at(arrow.from.row),
      x2: centerX.at(arrow.to.column),
      y2: centerY.at(arrow.to.row),
    }
    const key = `${idPrefix}-arrow-${index}`
    disposables.push(runtime.univerAPI.registerComponent(key, () => <ArrowGlyph {...geometry} />))
    const floating = worksheet.addFloatDomToRange(
      worksheet.getRange(
        box.startRow,
        box.startColumn,
        box.endRow - box.startRow + 1,
        box.endColumn - box.startColumn + 1,
      ),
      { componentKey: key, allowTransform: false, eventPassThrough: true },
      {},
      key,
    )
    if (floating) disposables.push(floating)
  })
  return disposables
}

function axisCenters(
  start: number,
  end: number,
  sizeOf: (index: number) => number,
): { total: number; at: (index: number) => number } {
  const offsets = new Map<number, number>()
  let total = 0
  for (let index = start; index <= end; index += 1) {
    const size = Math.max(sizeOf(index), 1)
    offsets.set(index, total + size / 2)
    total += size
  }
  return { total, at: (index) => offsets.get(index) ?? 0 }
}

function ArrowGlyph({
  width,
  height,
  x1,
  y1,
  x2,
  y2,
}: {
  readonly width: number
  readonly height: number
  readonly x1: number
  readonly y1: number
  readonly x2: number
  readonly y2: number
}): React.JSX.Element {
  const length = Math.hypot(x2 - x1, y2 - y1) || 1
  const ux = (x2 - x1) / length
  const uy = (y2 - y1) / length
  const head = Math.min(9, length / 2)
  const half = head * 0.42
  const baseX = x2 - ux * head
  const baseY = y2 - uy * head
  return (
    <svg
      className="trace-arrow-layer"
      viewBox={`0 0 ${width} ${height}`}
      preserveAspectRatio="none"
      aria-hidden="true"
    >
      <circle cx={x1} cy={y1} r={3} fill={ARROW_COLOR} />
      <line x1={x1} y1={y1} x2={baseX} y2={baseY} stroke={ARROW_COLOR} strokeWidth={1.5} />
      <polygon
        fill={ARROW_COLOR}
        points={
          `${x2},${y2} ` +
          `${baseX - uy * half},${baseY + ux * half} ` +
          `${baseX + uy * half},${baseY - ux * half}`
        }
      />
    </svg>
  )
}
