/**
 * Excel's outline gutter: a strip left of the row headers and one above the
 * column headers, drawn on DOM canvases beside the Univer container (the
 * grid shrinks to make room). Geometry comes from the live skeleton and
 * viewports, so scroll, zoom and frozen panes line up by construction.
 */
import type { IScale } from '@univerjs/core'
import {
  IRenderManagerService,
  SHEET_VIEWPORT_KEY,
  SheetColumnHeaderExtensionRegistry,
  SheetExtension,
  type SpreadsheetSkeleton,
  type UniverRenderingContext,
} from '@univerjs/engine-render'
import { SheetSkeletonManagerService } from '@univerjs/preset-sheets-core'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'

import { t } from './i18n/locale'
import {
  outlineColumnCenter,
  outlineGutterThickness,
  outlineMaxLevel,
  OUTLINE_GUTTER_PAD,
  OUTLINE_LEVEL_STEP,
  type OutlineGroup,
} from './outline-model'
import { outlineGroupsFor, type OutlineAxis } from './outline-actions'
import { sheetOutline } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'

export interface OutlineGutterDeps {
  readonly runtime: UniverRuntime
  readonly host: HTMLElement
  readonly state: () => LazyWorkbookState | null
  readonly onToggleGroup: (axis: OutlineAxis, group: OutlineGroup, collapsed: boolean) => void
  readonly onLevel: (axis: OutlineAxis, level: number) => void
}

export interface OutlineGutterHandle {
  refresh(): void
  dispose(): void
}

interface Hit {
  readonly x: number
  readonly y: number
  readonly w: number
  readonly h: number
  readonly action: () => void
}

interface Viewport {
  readonly isActive: boolean
  readonly top: number
  readonly left: number
  readonly width: number
  readonly height: number
  readonly viewportScrollX: number
  readonly viewportScrollY: number
}

interface Palette {
  readonly surface: string
  readonly line: string
  readonly border: string
  readonly text: string
}

const BUTTON_SIZE = OUTLINE_LEVEL_STEP - 2
/// Sheet id → height of the column-gutter band carved out of the header.
const columnGutterBand = new Map<string, number>()

interface HeaderStyleSource {
  uKey: string
  getHeaderStyle(sheetId: string): {
    fontSize?: number
    fontFamily?: string
    fontColor?: string
    backgroundColor?: string
    borderColor?: string
  }
}

/// Univer centers the column letters in the whole header; with the gutter
/// band on top they would sit under it, so repaint the band below it.
class OutlineColumnHeaderExtension extends SheetExtension {
  override uKey = 'GenOfficeOutlineColumnHeaderExtension'
  protected override Z_INDEX = 11

  override draw(
    ctx: UniverRenderingContext,
    parentScale: IScale,
    skeleton: SpreadsheetSkeleton,
  ): void {
    const band = columnGutterBand.get(skeleton.worksheet.getSheetId()) ?? 0
    const { columnHeaderHeight, columnWidthAccumulation, columnTotalWidth, rowColumnSegment } =
      skeleton
    if (band <= 0 || band >= columnHeaderHeight) return
    const source = (SheetColumnHeaderExtensionRegistry.getData() as HeaderStyleSource[]).find(
      (extension) => extension.uKey === 'DefaultColumnHeaderLayoutExtension',
    )
    const style = source?.getHeaderStyle(skeleton.worksheet.getSheetId()) ?? {}
    const scale = this._getScale(parentScale)
    const height = columnHeaderHeight - band
    ctx.save()
    ctx.translate(0, band)
    ctx.fillStyle = style.backgroundColor ?? ''
    ctx.fillRect(0, 0, columnTotalWidth, height)
    ctx.strokeStyle = style.borderColor ?? ''
    ctx.fillStyle = style.fontColor ?? ''
    ctx.font = `${style.fontSize ?? 13}px ${style.fontFamily ?? 'sans-serif'}`
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.lineWidth = 1 / scale
    const offset = 0.5 / scale
    let left =
      rowColumnSegment.startColumn > 0
        ? (columnWidthAccumulation[rowColumnSegment.startColumn - 1] ?? 0)
        : 0
    for (
      let column = rowColumnSegment.startColumn;
      column <= rowColumnSegment.endColumn;
      column += 1
    ) {
      const right = columnWidthAccumulation[column]
      if (right === undefined) break
      if (right > left) {
        ctx.beginPath()
        ctx.moveTo(right - offset, 0)
        ctx.lineTo(right - offset, height)
        ctx.stroke()
        ctx.fillText(columnLabel(column), (left + right) / 2, height / 2 + 1)
      }
      left = right
    }
    ctx.beginPath()
    ctx.moveTo(0, height - offset)
    ctx.lineTo(columnTotalWidth, height - offset)
    ctx.stroke()
    ctx.restore()
  }
}

SheetColumnHeaderExtensionRegistry.add(new OutlineColumnHeaderExtension())
const MIN_STRIP = OUTLINE_LEVEL_STEP + OUTLINE_GUTTER_PAD * 2
/// Univer's default column header height; the column gutter is carved out of
/// a taller header because the formula bar sits directly above it.
const DEFAULT_COLUMN_HEADER_HEIGHT = 20

/// Index of the first line whose end edge is past `offset` (accumulation is
/// cumulative end positions).
function firstLineAt(accumulation: readonly number[], offset: number): number {
  let low = 0
  let high = accumulation.length - 1
  while (low < high) {
    const mid = (low + high) >> 1
    if ((accumulation[mid] ?? 0) <= offset) low = mid + 1
    else high = mid
  }
  return low
}

function cssVar(element: HTMLElement, name: string): string {
  return getComputedStyle(element).getPropertyValue(name).trim()
}

export function installOutlineGutter(deps: OutlineGutterDeps): OutlineGutterHandle {
  const { runtime, host } = deps
  const rowCanvas = document.createElement('canvas')
  rowCanvas.className = 'outline-gutter outline-gutter-rows'
  const colCanvas = document.createElement('canvas')
  colCanvas.className = 'outline-gutter outline-gutter-cols'
  rowCanvas.setAttribute('aria-label', t('appGroupOutline'))
  colCanvas.setAttribute('aria-label', t('appGroupOutline'))
  let hits: { rows: Hit[]; cols: Hit[] } = { rows: [], cols: [] }
  let signature = ''
  let shown = false
  let disposed = false
  let frame = 0
  let groupCache: {
    key: string
    rows: OutlineGroup[]
    cols: OutlineGroup[]
  } | null = null
  let headerRequest: Promise<unknown> | null = null

  const hitAt = (axis: OutlineAxis, event: MouseEvent): Hit | undefined => {
    const canvas = axis === 'rows' ? rowCanvas : colCanvas
    const bounds = canvas.getBoundingClientRect()
    const x = event.clientX - bounds.left
    const y = event.clientY - bounds.top
    return hits[axis].find(
      (hit) => x >= hit.x && x <= hit.x + hit.w && y >= hit.y && y <= hit.y + hit.h,
    )
  }
  for (const axis of ['rows', 'cols'] as const) {
    const canvas = axis === 'rows' ? rowCanvas : colCanvas
    canvas.addEventListener('click', (event) => {
      const hit = hitAt(axis, event)
      if (!hit) return
      event.preventDefault()
      hit.action()
      signature = ''
    })
    canvas.addEventListener('mousemove', (event) => {
      canvas.style.cursor = hitAt(axis, event) ? 'pointer' : ''
    })
  }

  // Mounted only while an outline exists, after the grid: nothing that looks
  // for "the sheet canvas" (tests, tooling) should meet the gutter first.
  const setShown = (next: boolean, width: number): void => {
    host.classList.toggle('has-outline', next)
    host.style.setProperty('--outline-gutter-w', `${width}px`)
    if (next && !rowCanvas.isConnected) host.append(rowCanvas, colCanvas)
    if (!next && rowCanvas.isConnected) {
      rowCanvas.remove()
      colCanvas.remove()
    }
    shown = next
  }

  /// Grow/restore the column header so the column gutter has its band.
  const ensureHeaderHeight = (
    worksheet: UniverWorksheet,
    skeleton: { columnHeaderHeight: number },
    unitId: string,
    extra: number,
  ): boolean => {
    const current = skeleton.columnHeaderHeight
    const wanted = DEFAULT_COLUMN_HEADER_HEIGHT + extra
    if (current === wanted) return true
    if (current === 0 || headerRequest) return false
    headerRequest = runtime.univerAPI
      .executeCommand('sheet.command.set-col-header-height', {
        unitId,
        subUnitId: worksheet.getSheetId(),
        size: wanted,
      })
      .finally(() => {
        headerRequest = null
      })
    return false
  }

  const resolveTarget = () => {
    const state = deps.state()
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const worksheet = workbook?.getActiveSheet()
    if (!state || !workbook || !worksheet) return null
    const render = runtime.univer
      .__getInjector()
      .get(IRenderManagerService)
      .getRenderById(workbook.getId())
    const skeleton = render?.with(SheetSkeletonManagerService).getCurrentSkeleton()
    if (!render || !skeleton) return null
    return { state, worksheet, sheetId: worksheet.getSheetId(), render, skeleton }
  }

  const palette = (): Palette => ({
    surface: cssVar(host, '--surface-subtle'),
    line: cssVar(host, '--text-secondary'),
    border: cssVar(host, '--border-strong'),
    text: cssVar(host, '--text'),
  })

  const prepare = (
    canvas: HTMLCanvasElement,
    width: number,
    height: number,
  ): CanvasRenderingContext2D | null => {
    const dpr = window.devicePixelRatio || 1
    const pixelWidth = Math.max(1, Math.round(width * dpr))
    const pixelHeight = Math.max(1, Math.round(height * dpr))
    if (canvas.width !== pixelWidth) canvas.width = pixelWidth
    if (canvas.height !== pixelHeight) canvas.height = pixelHeight
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    ctx.clearRect(0, 0, width, height)
    return ctx
  }

  const drawButton = (
    ctx: CanvasRenderingContext2D,
    colors: Palette,
    cx: number,
    cy: number,
    glyph: string,
  ): { x: number; y: number; w: number; h: number } => {
    const x = Math.round(cx - BUTTON_SIZE / 2) + 0.5
    const y = Math.round(cy - BUTTON_SIZE / 2) + 0.5
    ctx.fillStyle = colors.surface
    ctx.strokeStyle = colors.border
    ctx.lineWidth = 1
    ctx.fillRect(x, y, BUTTON_SIZE, BUTTON_SIZE)
    ctx.strokeRect(x, y, BUTTON_SIZE, BUTTON_SIZE)
    ctx.fillStyle = colors.text
    ctx.font = `600 ${BUTTON_SIZE - 2}px var(--gs-font-sans, sans-serif)`
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.fillText(glyph, x + BUTTON_SIZE / 2, y + BUTTON_SIZE / 2 + 0.5)
    return { x: x - 1, y: y - 1, w: BUTTON_SIZE + 2, h: BUTTON_SIZE + 2 }
  }

  const drawLevelButtons = (
    ctx: CanvasRenderingContext2D,
    colors: Palette,
    axis: OutlineAxis,
    maxLevel: number,
    horizontal: boolean,
    into: Hit[],
  ): void => {
    for (let level = 1; level <= maxLevel + 1; level += 1) {
      const along = outlineColumnCenter(level)
      const across = OUTLINE_GUTTER_PAD + OUTLINE_LEVEL_STEP / 2
      const rect = horizontal
        ? drawButton(ctx, colors, along, across, String(level))
        : drawButton(ctx, colors, across, along, String(level))
      into.push({ ...rect, action: () => deps.onLevel(axis, level) })
    }
  }

  /// Paints one axis: `along` is the scrolling direction (y for rows).
  const drawAxis = (
    ctx: CanvasRenderingContext2D,
    colors: Palette,
    axis: OutlineAxis,
    groups: readonly OutlineGroup[],
    entries: ReadonlyMap<number, { level: number }>,
    maxLevel: number,
    viewports: readonly Viewport[],
    accumulation: readonly number[],
    scale: number,
    canvasOffset: number,
    into: Hit[],
    summaryAfter: boolean,
  ): void => {
    const vertical = axis === 'rows'
    const lineStart = (index: number): number => (index > 0 ? (accumulation[index - 1] ?? 0) : 0)
    const lineEnd = (index: number): number => accumulation[index] ?? lineStart(index)
    for (const viewport of viewports) {
      if (!viewport.isActive || (vertical ? viewport.height : viewport.width) <= 0) continue
      const viewStart = (vertical ? viewport.top : viewport.left) + canvasOffset
      const viewSize = vertical ? viewport.height : viewport.width
      const scroll = vertical ? viewport.viewportScrollY : viewport.viewportScrollX
      const toCanvas = (content: number): number => viewStart + (content - scroll) * scale
      const first = firstLineAt(accumulation, scroll)
      const last = Math.min(
        accumulation.length - 1,
        firstLineAt(accumulation, scroll + viewSize / scale),
      )
      ctx.save()
      ctx.beginPath()
      if (vertical) ctx.rect(0, viewStart, outlineGutterThickness(maxLevel) + MIN_STRIP, viewSize)
      else ctx.rect(viewStart, 0, viewSize, outlineGutterThickness(maxLevel) + MIN_STRIP)
      ctx.clip()
      ctx.strokeStyle = colors.line
      ctx.fillStyle = colors.line
      ctx.lineWidth = 1
      // Detail dots: one per visible grouped line, in the column past its level.
      for (let index = first; index <= last; index += 1) {
        const level = entries.get(index)?.level ?? 0
        if (level === 0 || lineEnd(index) <= lineStart(index)) continue
        const mid = toCanvas((lineStart(index) + lineEnd(index)) / 2)
        const across = outlineColumnCenter(level + 1)
        ctx.beginPath()
        if (vertical) ctx.arc(across, mid, 1.5, 0, Math.PI * 2)
        else ctx.arc(mid, across, 1.5, 0, Math.PI * 2)
        ctx.fill()
      }
      for (const group of groups) {
        const across = Math.round(outlineColumnCenter(group.level)) + 0.5
        if (!group.collapsed && group.end >= first && group.start <= last) {
          const from = toCanvas(lineStart(group.start))
          const to = toCanvas(lineEnd(group.end))
          ctx.beginPath()
          if (vertical) {
            ctx.moveTo(across, Math.round(from) + 0.5)
            ctx.lineTo(across, Math.round(to) - 0.5)
            const tick = summaryAfter ? Math.round(from) + 0.5 : Math.round(to) - 0.5
            ctx.moveTo(across, tick)
            ctx.lineTo(across + OUTLINE_LEVEL_STEP / 2, tick)
          } else {
            ctx.moveTo(Math.round(from) + 0.5, across)
            ctx.lineTo(Math.round(to) - 0.5, across)
            const tick = summaryAfter ? Math.round(from) + 0.5 : Math.round(to) - 0.5
            ctx.moveTo(tick, across)
            ctx.lineTo(tick, across + OUTLINE_LEVEL_STEP / 2)
          }
          ctx.stroke()
        }
        if (group.summary < first || group.summary > last) continue
        if (lineEnd(group.summary) <= lineStart(group.summary)) continue
        const mid = toCanvas((lineStart(group.summary) + lineEnd(group.summary)) / 2)
        const rect = vertical
          ? drawButton(ctx, colors, across, mid, group.collapsed ? '+' : '−')
          : drawButton(ctx, colors, mid, across, group.collapsed ? '+' : '−')
        into.push({
          ...rect,
          action: () => deps.onToggleGroup(axis, group, !group.collapsed),
        })
      }
      ctx.restore()
    }
  }

  const paint = (): void => {
    const target = resolveTarget()
    if (!target) {
      if (shown) setShown(false, 0)
      return
    }
    const { state, worksheet, sheetId, render, skeleton } = target
    const outline = sheetOutline(state, sheetId)
    const rowDepth = outlineMaxLevel(outline.rows)
    const colDepth = outlineMaxLevel(outline.cols)
    const unitId = runtime.univerAPI.getActiveWorkbook()?.getId() ?? ''
    if (rowDepth === 0 && colDepth === 0) {
      if (shown) setShown(false, 0)
      signature = ''
      ensureHeaderHeight(worksheet, skeleton, unitId, 0)
      if (columnGutterBand.get(sheetId)) {
        columnGutterBand.set(sheetId, 0)
        render.scene.makeDirty(true)
      }
      return
    }
    const scene = render.scene
    const { scaleX, scaleY } = scene.getAncestorScale()
    const viewMain = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN) as Viewport | undefined
    const viewTop = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN_TOP) as Viewport | undefined
    const viewLeft = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN_LEFT) as Viewport | undefined
    if (!viewMain) return
    const width = outlineGutterThickness(rowDepth)
    const height = outlineGutterThickness(colDepth)
    if (!shown || host.style.getPropertyValue('--outline-gutter-w') !== `${width}px`) {
      setShown(true, width)
    }
    const headerReady = ensureHeaderHeight(worksheet, skeleton, unitId, height)
    const band = headerReady ? height : 0
    if (columnGutterBand.get(sheetId) !== band) {
      columnGutterBand.set(sheetId, band)
      render.scene.makeDirty(true)
    }
    const gridCanvas = render.engine.getCanvasElement()
    const gridBounds = gridCanvas.getBoundingClientRect()
    const hostBounds = host.getBoundingClientRect()
    colCanvas.style.left = `${gridBounds.left - hostBounds.left}px`
    colCanvas.style.top = `${gridBounds.top - hostBounds.top}px`
    colCanvas.style.width = `${gridBounds.width}px`
    colCanvas.style.height = `${headerReady ? height : 0}px`
    const rowBounds = rowCanvas.getBoundingClientRect()
    const colBounds = colCanvas.getBoundingClientRect()
    const next = [
      sheetId,
      outline.version,
      skeleton.rowTotalHeight,
      skeleton.columnTotalWidth,
      scaleX,
      scaleY,
      viewMain.viewportScrollX,
      viewMain.viewportScrollY,
      viewMain.top,
      viewMain.left,
      viewTop?.height ?? 0,
      viewLeft?.width ?? 0,
      gridBounds.top - rowBounds.top,
      rowBounds.height,
      colBounds.width,
      colBounds.height,
      skeleton.columnHeaderHeight,
      host.dataset['theme'] ?? document.documentElement.getAttribute('data-theme') ?? '',
    ].join('|')
    if (next === signature) return
    signature = next
    const cacheKey = `${sheetId}|${outline.version}|${skeleton.rowTotalHeight}|${skeleton.columnTotalWidth}`
    if (groupCache?.key !== cacheKey) {
      groupCache = {
        key: cacheKey,
        rows: rowDepth > 0 ? outlineGroupsFor(state, worksheet, sheetId, 'rows') : [],
        cols: colDepth > 0 ? outlineGroupsFor(state, worksheet, sheetId, 'cols') : [],
      }
    }
    const colors = palette()
    const nextHits: { rows: Hit[]; cols: Hit[] } = { rows: [], cols: [] }

    const gridTop = gridBounds.top - rowBounds.top
    const rowCtx = prepare(rowCanvas, Math.max(width, 1), rowBounds.height)
    if (rowCtx && width > 0) {
      rowCtx.fillStyle = colors.surface
      rowCtx.fillRect(0, gridTop, width, rowBounds.height - gridTop)
      rowCtx.fillStyle = colors.border
      rowCtx.fillRect(width - 1, gridTop, 1, rowBounds.height - gridTop)
      // Row level buttons sit beside the column header, as in Excel.
      const stripTop = gridTop + Math.max(0, (viewMain.top - MIN_STRIP) / 2)
      rowCtx.save()
      rowCtx.translate(0, stripTop)
      const levelHits: Hit[] = []
      drawLevelButtons(rowCtx, colors, 'rows', rowDepth, true, levelHits)
      rowCtx.restore()
      for (const hit of levelHits) nextHits.rows.push({ ...hit, y: hit.y + stripTop })
      drawAxis(
        rowCtx,
        colors,
        'rows',
        groupCache.rows,
        outline.rows,
        rowDepth,
        [viewTop, viewMain].filter((viewport): viewport is Viewport => Boolean(viewport)),
        skeleton.rowHeightAccumulation,
        scaleY,
        gridTop,
        nextHits.rows,
        outline.summaryBelow,
      )
    }
    const colCtx = prepare(colCanvas, Math.max(colBounds.width, 1), Math.max(height, 1))
    if (colCtx && height > 0 && colBounds.height > 0) {
      colCtx.fillStyle = colors.surface
      colCtx.fillRect(0, 0, colBounds.width, height)
      colCtx.fillStyle = colors.border
      colCtx.fillRect(0, height - 1, colBounds.width, 1)
      {
        drawLevelButtons(colCtx, colors, 'cols', colDepth, false, nextHits.cols)
        drawAxis(
          colCtx,
          colors,
          'cols',
          groupCache.cols,
          outline.cols,
          colDepth,
          [viewLeft, viewMain].filter((viewport): viewport is Viewport => Boolean(viewport)),
          skeleton.columnWidthAccumulation,
          scaleX,
          0,
          nextHits.cols,
          outline.summaryRight,
        )
      }
    }
    hits = nextHits
  }

  const tick = (): void => {
    if (disposed) return
    try {
      paint()
    } catch (error) {
      // Render modules can be mid-rebuild while a workbook swaps in.
      signature = ''
      console.warn('[outline-gutter]', error)
    }
    frame = window.requestAnimationFrame(tick)
  }
  frame = window.requestAnimationFrame(tick)

  return {
    refresh() {
      signature = ''
      groupCache = null
    },
    dispose() {
      disposed = true
      window.cancelAnimationFrame(frame)
      rowCanvas.remove()
      colCanvas.remove()
      columnGutterBand.clear()
      setShown(false, 0)
    },
  }
}
