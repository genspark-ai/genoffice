/**
 * Cross-highlight ("reading mode"): a lightweight canvas extension paints
 * the active cell's visible row and column. Drawing inside the grid avoids
 * float-DOM observer leaks, never sits above chart/shape DOM, and naturally
 * works in blank cells and sheets larger than the used-data extent.
 */
import type { IRange, IScale } from '@univerjs/core'
import {
  IRenderManagerService,
  SheetExtension,
  SpreadsheetExtensionRegistry,
  type IDrawInfo,
  type SpreadsheetSkeleton,
  type UniverRenderingContext,
} from '@univerjs/engine-render'
import type { UniverRuntime } from './univer-state'

/** Canvas chrome colors, kept out of document data and keyed by UI theme. */
export const CROSS_HIGHLIGHT_CANVAS_COLORS = {
  light: {
    fill: 'rgba(31, 90, 168, 0.07)',
    line: 'rgba(31, 90, 168, 0.30)',
  },
  dark: {
    fill: 'rgba(96, 165, 250, 0.12)',
    line: 'rgba(96, 165, 250, 0.35)',
  },
} as const

export type CrossHighlightTheme = keyof typeof CROSS_HIGHLIGHT_CANVAS_COLORS

interface CellRect {
  readonly startX: number
  readonly startY: number
  readonly endX: number
  readonly endY: number
}

export interface CrossHighlightRect {
  readonly key: 'row' | 'column'
  readonly left: number
  readonly top: number
  readonly width: number
  readonly height: number
}

/**
 * Rectangles needed in the current viewport(s). Only visible cells are
 * inspected, so empty sheets and rows beyond the old 20k cap cost the same
 * two rectangles as a small populated sheet.
 */
export function crossHighlightRects(
  activeRow: number,
  activeColumn: number,
  viewRanges: readonly IRange[],
  cellAt: (row: number, column: number) => CellRect | null,
): CrossHighlightRect[] {
  if (
    !Number.isInteger(activeRow) ||
    !Number.isInteger(activeColumn) ||
    activeRow < 0 ||
    activeColumn < 0
  ) {
    return []
  }
  const rects: CrossHighlightRect[] = []
  const seen = new Set<string>()
  const push = (rect: CrossHighlightRect): void => {
    if (!(rect.width > 0) || !(rect.height > 0)) return
    const key = `${rect.key}:${rect.left}:${rect.top}:${rect.width}:${rect.height}`
    if (seen.has(key)) return
    seen.add(key)
    rects.push(rect)
  }
  for (const range of viewRanges) {
    const startRow = Math.max(0, Math.floor(range.startRow))
    const endRow = Math.max(startRow, Math.floor(range.endRow))
    const startColumn = Math.max(0, Math.floor(range.startColumn))
    const endColumn = Math.max(startColumn, Math.floor(range.endColumn))
    if (activeRow >= startRow && activeRow <= endRow) {
      const first = cellAt(activeRow, startColumn)
      const last = cellAt(activeRow, endColumn)
      if (first && last) {
        push({
          key: 'row',
          left: first.startX,
          top: first.startY,
          width: last.endX - first.startX,
          height: first.endY - first.startY,
        })
      }
    }
    if (activeColumn >= startColumn && activeColumn <= endColumn) {
      const first = cellAt(startRow, activeColumn)
      const last = cellAt(endRow, activeColumn)
      if (first && last) {
        push({
          key: 'column',
          left: first.startX,
          top: first.startY,
          width: first.endX - first.startX,
          height: last.endY - first.startY,
        })
      }
    }
  }
  return rects
}

const STORAGE_KEY = 'ai-sheets-cross-highlight'

/** The persisted View-tab toggle; defaults to off (also headless-safe). */
export function loadCrossHighlightPreference(): boolean {
  try {
    return window.localStorage.getItem(STORAGE_KEY) === '1'
  } catch {
    // No localStorage (tests, blocked storage): the safe default is off.
    return false
  }
}

export function storeCrossHighlightPreference(enabled: boolean): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, enabled ? '1' : '0')
  } catch {
    // Preference stays session-only when storage is unavailable.
  }
}

export interface CrossHighlightOptions {
  readonly theme: () => CrossHighlightTheme
}

export interface CrossHighlightHandle {
  setVisible(visible: boolean): void
  refresh(): void
  dispose(): void
}

export interface CrossHighlightTarget {
  readonly workbookId: string
  readonly sheetId: string
  readonly row: number
  readonly column: number
}

interface CrossHighlightWorkbook {
  getId(): string
  getActiveSheet(): { getSheetId(): string } | null
  getActiveCell(): { getRow(): number; getColumn(): number } | null
}

/** Resolve the selection's primary cell, not the range rectangle's top-left. */
export function resolveCrossHighlightTarget(
  workbook: CrossHighlightWorkbook | null,
): CrossHighlightTarget | null {
  if (!workbook) return null
  const worksheet = workbook.getActiveSheet()
  const cell = workbook.getActiveCell()
  if (!worksheet || !cell) return null
  const row = cell.getRow()
  const column = cell.getColumn()
  if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0) return null
  return {
    workbookId: workbook.getId(),
    sheetId: worksheet.getSheetId(),
    row,
    column,
  }
}

interface CrossHighlightRenderState extends CrossHighlightTarget {
  readonly visible: boolean
  readonly theme: CrossHighlightTheme
}

let renderState: CrossHighlightRenderState = {
  visible: false,
  workbookId: '',
  sheetId: '',
  row: -1,
  column: -1,
  theme: 'light',
}

/**
 * Grid extension ordered just above authored cell backgrounds and below
 * fonts/borders. It renders no document content and is disabled for print.
 */
export class CrossHighlightExtension extends SheetExtension {
  readonly uKey = 'GenOfficeCrossHighlightExtension'
  protected Z_INDEX = 22

  override draw(
    ctx: UniverRenderingContext,
    _parentScale: IScale,
    skeleton: SpreadsheetSkeleton,
    _diffRanges: IRange[] | undefined,
    drawInfo?: IDrawInfo,
  ): void {
    if (!renderState.visible || ctx.__mode === 'printing') return
    const worksheet = skeleton.worksheet
    if (
      !worksheet ||
      worksheet.unitId !== renderState.workbookId ||
      worksheet.getSheetId() !== renderState.sheetId
    ) {
      return
    }
    const viewRanges = drawInfo?.viewRanges?.length
      ? drawInfo.viewRanges
      : [skeleton.rowColumnSegment]
    const rects = crossHighlightRects(
      renderState.row,
      renderState.column,
      viewRanges,
      (row, column) => skeleton.getCellWithCoordByIndex(row, column, false),
    )
    if (rects.length === 0) return
    const colors = CROSS_HIGHLIGHT_CANVAS_COLORS[renderState.theme]
    ctx.save()
    ctx.fillStyle = colors.fill
    for (const rect of rects) ctx.fillRect(rect.left, rect.top, rect.width, rect.height)
    const { scaleX, scaleY } = ctx.getScale()
    ctx.strokeStyle = colors.line
    ctx.lineWidth = 1 / Math.max(scaleX, scaleY, 1)
    ctx.beginPath()
    for (const rect of rects) {
      if (rect.key === 'row') {
        ctx.moveTo(rect.left, rect.top + rect.height)
        ctx.lineTo(rect.left + rect.width, rect.top + rect.height)
      } else {
        ctx.moveTo(rect.left + rect.width, rect.top)
        ctx.lineTo(rect.left + rect.width, rect.top + rect.height)
      }
    }
    ctx.stroke()
    ctx.closePath()
    ctx.restore()
  }
}

SpreadsheetExtensionRegistry.add(CrossHighlightExtension)

function markRenderDirty(runtime: UniverRuntime, workbookId: string): void {
  if (!workbookId) return
  try {
    const render = runtime.univer
      .__getInjector()
      .get(IRenderManagerService)
      .getRenderById(workbookId)
    render?.mainComponent?.makeDirty(true)
    render?.scene?.makeDirty(true)
  } catch {
    // Render modules may not exist yet during workbook replacement.
  }
}

/**
 * Keeps the shared canvas state aligned with the active workbook selection.
 * Selection moves update four numbers and repaint; no per-move resources are
 * allocated, so long navigation sessions cannot accumulate observers.
 */
export function installCrossHighlight(
  runtime: UniverRuntime,
  options: CrossHighlightOptions,
): CrossHighlightHandle {
  let visible = false
  let disposed = false

  const sync = (): void => {
    if (disposed) return
    const previousWorkbookId = renderState.visible ? renderState.workbookId : ''
    const target = (() => {
      try {
        return visible ? resolveCrossHighlightTarget(runtime.univerAPI.getActiveWorkbook()) : null
      } catch {
        return null
      }
    })()
    const next: CrossHighlightRenderState = target
      ? { ...target, visible: true, theme: options.theme() }
      : {
          visible: false,
          workbookId: '',
          sheetId: '',
          row: -1,
          column: -1,
          theme: options.theme(),
        }
    if (
      next.visible === renderState.visible &&
      next.workbookId === renderState.workbookId &&
      next.sheetId === renderState.sheetId &&
      next.row === renderState.row &&
      next.column === renderState.column &&
      next.theme === renderState.theme
    ) {
      return
    }
    renderState = next
    markRenderDirty(runtime, previousWorkbookId)
    if (next.workbookId !== previousWorkbookId) markRenderDirty(runtime, next.workbookId)
  }

  const disposables = [
    runtime.univerAPI.addEvent(runtime.univerAPI.Event.SelectionChanged, sync),
    runtime.univerAPI.addEvent(runtime.univerAPI.Event.ActiveSheetChanged, sync),
  ]

  return {
    setVisible(next: boolean): void {
      visible = next
      sync()
    },
    refresh(): void {
      sync()
    },
    dispose(): void {
      visible = false
      sync()
      disposed = true
      for (const disposable of disposables) disposable.dispose()
    },
  }
}
