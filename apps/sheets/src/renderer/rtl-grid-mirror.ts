/**
 * Whole-grid mirroring for Excel's `sheetView rightToLeft="1"`: column A at
 * the right edge, columns running right-to-left, the row-header strip on the
 * right. Univer 0.25.1 stores `rightToLeft` on the worksheet snapshot but no
 * renderer code reads it, so this module supplies the rendering.
 *
 * Approach: mirror GEOMETRY, keep LOGICAL indices. Every paint and hit-test
 * path funnels through a handful of skeleton methods
 * (`getCellWithCoordByIndex` for cells/borders/fonts/selection,
 * `getColumnIndexByOffsetX` for pointer hits, `_getRangeByViewBounding` for
 * visible-range culling and lazy loading, `getOffsetRelativeToRowCol` +
 * `getCellWithCoordByIndex` for the scroll-state round trip), so patching
 * those prototypes mirrors content, selection, editing and scrolling in one
 * place. A canvas `scale(-1, 1)` flip is used only for text-free paint code
 * (gridlines, header cell frames); text always renders unflipped.
 *
 * Column freeze is mirrored too: the frozen pane docks against the right-edge
 * row-header strip and the scrollable pane takes the band to its left. The
 * stock freeze layout (`_updateViewport`) does its gap/padding math from
 * `getNoMergeCellWithCoordByIndex`, which the patches above mirror into scene
 * space while that math needs LTR data space — so the stock pass runs with
 * the mirror suspended and the viewports are re-anchored afterwards.
 *
 * Scrolling semantics: the RTL "home" (column A flush right) is the MAX
 * viewportScrollX, recorded as the sentinel state column 0 / offset 0 so
 * resizes re-anchor flush-right. Everything that derives a scroll target from
 * that state in LTR terms is re-derived in scene space here: relative (wheel)
 * scrolls, the minimal reveal of an off-screen cell, and the column freeze
 * divider drag.
 */
import {
  BooleanNumber,
  createInterceptorKey,
  ICommandService,
  IUniverInstanceService,
  WrapStrategy,
} from '@univerjs/core'
import type { IRange, Nullable } from '@univerjs/core'
import {
  Border,
  ColumnHeaderLayout,
  CURSOR_TYPE,
  Font,
  getDocsSkeletonPageSize,
  IRenderManagerService,
  SHEET_VIEWPORT_KEY,
  Spreadsheet,
  SpreadsheetColumnHeader,
  SpreadsheetRowHeader,
  SpreadsheetSkeleton,
  Viewport,
} from '@univerjs/engine-render'
import { getSheetCommandTarget, SetFrozenCommand } from '@univerjs/sheets'
import {
  getCoordByOffset,
  HeaderFreezeRenderController,
  ScrollCommand,
  SetScrollOperation,
  SetScrollRelativeCommand,
  SheetSkeletonManagerService,
  SheetsScrollRenderController,
} from '@univerjs/sheets-ui'

/** Scenes whose active sheet is RTL — consulted by shared prototype patches. */
const rtlScenes = new WeakSet<object>()

/** Scene → scroll manager, so viewport patches can ask "is this sheet at home?". */
const sceneScrollManagers = new WeakMap<object, any>()

let reapplyingRtlMainScroll = false

/** While true the skeleton patches report LTR (data-space) geometry. */
let rtlMirrorSuspended = false

/** Controller → sheetIds whose poisoned load-time X-state was already reset. */
const anchoredRtlSheets = new WeakMap<object, Set<string>>()

function isAtHorizontalHome(scene: object): boolean {
  const state = sceneScrollManagers.get(scene)?.getCurrentScrollState?.()
  return !state || ((state.sheetViewStartColumn ?? 0) === 0 && (state.offsetX ?? 0) === 0)
}

/** Logical-left borders paint on the visual right (and diagonals flip). */
export const RTL_BORDER_TYPE_SWAP: Record<string, string> = {
  l: 'r',
  r: 'l',
  tl_br: 'bl_tr',
  bl_tr: 'tl_br',
  tl_bc: 'bc_tr',
  bc_tr: 'tl_bc',
  tl_mr: 'ml_tr',
  ml_tr: 'tl_mr',
}

/** Mirror an x span inside the grid; `headerOffset` when coords include it. */
export function mirrorSpanX(
  startX: number,
  endX: number,
  totalWidth: number,
  headerOffset = 0,
): { startX: number; endX: number } {
  return {
    startX: headerOffset + totalWidth - (endX - headerOffset),
    endX: headerOffset + totalWidth - (startX - headerOffset),
  }
}

/**
 * Canvas placement for an RTL sheet's viewports (scene units unless noted).
 * `freezeStartX`/`freezeEndX` are the LTR (data-space) x of the freeze
 * anchor's pane start and end columns; both 0 when no column freeze.
 */
export function rtlFreezeLayout(params: {
  engineWidth: number
  scaleX: number
  headerWidth: number
  totalWidth: number
  freezeStartX: number
  freezeEndX: number
}): {
  stripLeft: number
  freezeGap: number
  paneLeft: number
  /** Right inset for the scrollable pane and its header strip, device px. */
  mainRightInset: number
  /** viewportScrollX pinning the mirrored frozen band at `paneLeft`. */
  frozenScrollX: number
} {
  const { engineWidth, scaleX, headerWidth, totalWidth, freezeStartX, freezeEndX } = params
  const stripLeft = engineWidth / scaleX - headerWidth - 1
  const freezeGap = Math.max(0, freezeEndX - freezeStartX)
  return {
    stripLeft,
    freezeGap,
    paneLeft: stripLeft - freezeGap,
    mainRightInset: (headerWidth + 1 + freezeGap) * scaleX,
    frozenScrollX: headerWidth + totalWidth - freezeStartX - stripLeft,
  }
}

/**
 * Highest useful viewportScrollX for the RTL main pane: grid right edge (or
 * the frozen band's left edge) flush with the pane's right edge. The pane's
 * padding is the RTL form set by patchFreezeViewportAnchors — [0, bandEnd]
 * in header-less LTR x, i.e. the excluded width is the frozen band plus the
 * columns hidden behind it.
 */
export function rtlMaxViewportScrollX(
  sceneWidth: number,
  viewportWidth: number,
  scaleX: number,
  paddingStartX = 0,
  paddingEndX = 0,
): number {
  return sceneWidth - Math.max(0, paddingEndX - paddingStartX) - viewportWidth / scaleX
}

interface SkeletonLike {
  worksheet?: { getConfig?(): { rightToLeft?: BooleanNumber } } | null
  columnTotalWidth: number
  columnWidthAccumulation: number[]
  rowHeaderWidthAndMarginLeft: number
  columnHeaderHeightAndMarginTop: number
}

export function isRtlSkeleton(skeleton: SkeletonLike | null | undefined): boolean {
  if (rtlMirrorSuspended) return false
  try {
    return skeleton?.worksheet?.getConfig?.().rightToLeft === BooleanNumber.TRUE
  } catch {
    return false
  }
}

interface CellCoordLike {
  startX: number
  endX: number
  startY: number
  endY: number
  mergeInfo?: { startX: number; endX: number; startY: number; endY: number }
}

interface ViewBoundLike {
  left: number
  right: number
  top: number
  bottom: number
}

interface RenderCtxLike {
  save(): void
  restore(): void
  translate(x: number, y: number): void
  scale(x: number, y: number): void
  fillText?: unknown
  rectByPrecision(x: number, y: number, w: number, h: number): void
  clip(): void
}

function mirrorCellCoord<T extends CellCoordLike>(
  cell: T,
  totalWidth: number,
  headerOffset: number,
): T {
  const span = mirrorSpanX(cell.startX, cell.endX, totalWidth, headerOffset)
  const mirrored = { ...cell, startX: span.startX, endX: span.endX }
  if (cell.mergeInfo) {
    const merge = mirrorSpanX(cell.mergeInfo.startX, cell.mergeInfo.endX, totalWidth, headerOffset)
    mirrored.mergeInfo = { ...cell.mergeInfo, startX: merge.startX, endX: merge.endX }
  }
  return mirrored
}

let installed = false

export function installRtlGridMirror(): void {
  if (installed) return
  installed = true
  patchSkeletonGeometry()
  patchSpreadsheetGridlines()
  patchBorderExtension()
  patchFontExtension()
  patchColumnHeader()
  patchRowHeader()
  patchViewportMargin()
  patchFreezeViewportAnchors()
  patchFreezeBars()
  patchRelativeScroll()
  patchRevealScroll()
  patchFreezeDrag()
}

/**
 * The column freeze divider is a scene Rect ending at the anchor cell's
 * startX (its left edge, LTR). Mirrored, the frozen band sits RIGHT of the
 * divider, so the bar belongs at the anchor cell's mirrored endX instead —
 * the stock placement lands one column into the scrollable pane.
 */
function patchFreezeBars(): void {
  const proto = HeaderFreezeRenderController.prototype as any
  const origCreate = proto._createFreeze
  proto._createFreeze = function (freezeDirectionType = 0, freezeConfig?: any) {
    const result = origCreate.call(this, freezeDirectionType, freezeConfig)
    if (freezeDirectionType === 0) return result
    const skeleton = this._sheetSkeletonManagerService?.getCurrentParam?.()?.skeleton
    const config = freezeConfig ?? this._getFreeze?.()
    // The zero-split ghost bar keeps its stock spot: the RTL drag path for
    // creating a freeze is not mirrored, so it stays out of sight as before.
    if (!isRtlSkeleton(skeleton) || !config || !(config.startColumn > 0)) return result
    // Row -1 (column-only freeze) would NaN out into the fallback position.
    const anchorRow = Math.max(0, config.startRow)
    const boundary = this._getPositionByIndex?.(anchorRow, config.startColumn)?.endX
    if (typeof boundary === 'number' && Number.isFinite(boundary)) {
      this._columnFreezeHeaderRect?.translate(boundary)
      this._columnFreezeMainRect?.translate(boundary)
    }
    return result
  }
}

/**
 * The horizontal scrollbar sizes its range from `scene.width - marginLeft`
 * (the grid area right of the header strip). With the RTL viewMain anchored
 * at x=0 the reachable range must include the header-width band too, or the
 * grid clamps `headerWidth` short of flush-right.
 */
function patchViewportMargin(): void {
  const proto = Viewport.prototype as any
  const origSetMargin = proto.setMargin
  proto.setMargin = function (marginLeft: number, marginTop: number) {
    const rtl = rtlScenes.has(this._scene) && this.viewportKey === SHEET_VIEWPORT_KEY.VIEW_MAIN
    origSetMargin.call(this, rtl ? 0 : marginLeft, marginTop)
  }

  // The RTL home position is the MAX scroll (column A flush right). The
  // scrollbar's own limit overshoots by its track slack, so clamp programmatic
  // scrolls to exactly "grid right edge flush at the viewport right edge".
  const origScrollToViewportPos = proto.scrollToViewportPos
  proto.scrollToViewportPos = function (
    pos: { viewportScrollX?: number; viewportScrollY?: number },
    isTrigger?: boolean,
  ) {
    if (
      rtlScenes.has(this._scene) &&
      this.viewportKey === SHEET_VIEWPORT_KEY.VIEW_MAIN &&
      typeof pos?.viewportScrollX === 'number'
    ) {
      const scaleX = this._scene.scaleX || 1
      const maxUseful = rtlMaxViewportScrollX(
        this._scene.width,
        this.width || 0,
        scaleX,
        this._paddingStartX,
        this._paddingEndX,
      )
      if (maxUseful > 0 && pos.viewportScrollX >= maxUseful) {
        // A request at or past the edge is the RTL "home" (state A1/0 maps
        // past the edge; wheel/reveal land exactly on it). Cap it, and keep
        // the clamped X out of the recorded scroll state (isTrigger=false)
        // so later restores re-cap against fresh geometry instead of
        // freezing a stale clamp.
        const result = origScrollToViewportPos.call(
          this,
          { ...pos, viewportScrollX: maxUseful },
          false,
        )
        // Untriggered scrolls bypass the stock header-strip sync.
        for (const key of [
          SHEET_VIEWPORT_KEY.VIEW_COLUMN_RIGHT,
          SHEET_VIEWPORT_KEY.VIEW_MAIN_TOP,
        ]) {
          const sibling = this._scene.getViewport(key)
          if (sibling) sibling.viewportScrollX = this.viewportScrollX
        }
        for (const key of [SHEET_VIEWPORT_KEY.VIEW_ROW_BOTTOM, SHEET_VIEWPORT_KEY.VIEW_MAIN_LEFT]) {
          const sibling = this._scene.getViewport(key)
          if (sibling) sibling.viewportScrollY = this.viewportScrollY
        }
        recordHomeScrollState(this)
        return result
      }
      if (reapplyingRtlMainScroll) isTrigger = false
    }
    return origScrollToViewportPos.call(this, pos, isTrigger)
  }

  // On resize, the stock re-apply replays the CURRENT pixel position and
  // records it as user scroll state — which freezes an RTL sheet still at
  // home onto a stale clamp. Replay without recording, then re-anchor
  // flush-right while the sheet is at home (Excel keeps column A pinned to
  // the right edge across resizes).
  const origUpdateScroll = proto._updateScrollByViewportScrollValue
  proto._updateScrollByViewportScrollValue = function () {
    const rtlMain = rtlScenes.has(this._scene) && this.viewportKey === SHEET_VIEWPORT_KEY.VIEW_MAIN
    if (!rtlMain) return origUpdateScroll.call(this)
    reapplyingRtlMainScroll = true
    try {
      origUpdateScroll.call(this)
    } finally {
      reapplyingRtlMainScroll = false
    }
    if (isAtHorizontalHome(this._scene)) {
      this.scrollToViewportPos({
        viewportScrollX: Number.MAX_SAFE_INTEGER,
        viewportScrollY: this.viewportScrollY,
      })
    }
  }
}

/**
 * The capped ("home") scroll skips the stock onScrollAfter$ bookkeeping, which
 * would otherwise freeze the clamped X into the scroll state. Record the home
 * sentinel (column 0 / offset 0) with the LIVE vertical position instead, and
 * publish it: relative scrolls read their baseline from this state, so a
 * stale Y here would pin vertical wheel scrolling while the sheet sits at its
 * horizontal home, and the sibling strips / lazy loader follow the publish.
 */
function recordHomeScrollState(viewport: any): void {
  const sm = sceneScrollManagers.get(viewport._scene)
  const skeleton = sm?._sheetSkeletonManagerService?.getCurrentSkeleton?.()
  if (!sm || !skeleton) return
  const { viewportScrollX, viewportScrollY, scrollX, scrollY } = viewport
  const { row, rowOffset } = skeleton.getOffsetRelativeToRowCol(viewportScrollX, viewportScrollY)
  const state = {
    sheetViewStartRow: row,
    sheetViewStartColumn: 0,
    offsetX: 0,
    offsetY: rowOffset,
    viewportScrollX,
    viewportScrollY,
    scrollX,
    scrollY,
  }
  sm.setValidScrollStateToCurrSheet?.(state)
  sm.validViewportScrollInfo$?.next?.(state)
}

interface RtlMainViewport {
  scene: any
  skeleton: any
  viewMain: any
  scaleX: number
  /** Scene-unit width of the scrollable pane. */
  paneWidth: number
  /** viewportScrollX range: [minScrollX, maxScrollX] (max = RTL home). */
  minScrollX: number
  maxScrollX: number
}

/**
 * The RTL main viewport of a sheet scene, or null when the sheet is not RTL.
 * The RTL viewMain is anchored at canvas x=0 (patchFreezeViewportAnchors), so
 * its visible scene window is exactly [viewportScrollX, viewportScrollX +
 * paneWidth].
 */
function rtlMainViewport(scene: any, skeleton: any): RtlMainViewport | null {
  if (!scene || !skeleton || !rtlScenes.has(scene) || !isRtlSkeleton(skeleton)) return null
  const viewMain = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
  if (!viewMain) return null
  const scaleX = scene.scaleX || 1
  const maxScrollX = rtlMaxViewportScrollX(
    scene.width,
    viewMain.width || 0,
    scaleX,
    viewMain._paddingStartX,
    viewMain._paddingEndX,
  )
  return {
    scene,
    skeleton,
    viewMain,
    scaleX,
    paneWidth: (viewMain.width || 0) / scaleX,
    // The highest columns start at scene x 0 (the RTL padding starts there).
    minScrollX: 0,
    maxScrollX,
  }
}

/**
 * Scroll-state fields for a target viewportScrollX on an RTL sheet: the
 * pixel round-trips through the mirrored getOffsetRelativeToRowCol, except
 * that a target at/past the home edge becomes the home sentinel (column 0 /
 * offset 0) so resize re-anchoring keeps working.
 */
export function rtlHorizontalScrollState(
  targetScrollX: number,
  bounds: { minScrollX: number; maxScrollX: number },
  offsetRelativeToColumn: (scrollX: number) => { column: number; columnOffset: number },
): { sheetViewStartColumn: number; offsetX: number } {
  if (bounds.maxScrollX <= 0 || targetScrollX >= bounds.maxScrollX) {
    return { sheetViewStartColumn: 0, offsetX: 0 }
  }
  const clamped = Math.max(bounds.minScrollX, targetScrollX)
  const { column, columnOffset } = offsetRelativeToColumn(clamped)
  return { sheetViewStartColumn: column, offsetX: columnOffset }
}

/**
 * Wheel / trackpad scrolling. The stock relative-scroll command adds the
 * delta to the RECORDED state, but the RTL home state is a sentinel whose
 * pixel lies past the edge, so every delta was re-capped to home and the
 * sheet never moved horizontally. Add the delta to the live viewport
 * position instead. No sign flip: scene x already grows toward the visual
 * right (column A sits at max x), so wheel-right lowers the column index
 * exactly as in Excel, and trackpad momentum is plain delta accumulation.
 */
function patchRelativeScroll(): void {
  const command = SetScrollRelativeCommand as unknown as {
    handler: (accessor: any, params?: { offsetX?: number; offsetY?: number }) => unknown
  }
  const origHandler = command.handler
  command.handler = (accessor, params) => {
    const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
    const render = target ? accessor.get(IRenderManagerService).getRenderById(target.unitId) : null
    const rtl = rtlMainViewport(
      render?.scene,
      render?.with(SheetSkeletonManagerService)?.getCurrentSkeleton(),
    )
    if (!target || !rtl) return origHandler(accessor, params)
    const { offsetX = 0, offsetY = 0 } = params ?? {}
    const { viewMain, skeleton } = rtl
    const targetX = viewMain.viewportScrollX + offsetX
    const targetY = Math.max(0, viewMain.viewportScrollY + offsetY)
    const { row, rowOffset } = skeleton.getOffsetRelativeToRowCol(0, targetY)
    const horizontal = rtlHorizontalScrollState(targetX, rtl, (scrollX) =>
      skeleton.getOffsetRelativeToRowCol(scrollX, 0),
    )
    return accessor.get(ICommandService).executeCommand(SetScrollOperation.id, {
      unitId: target.unitId,
      sheetId: target.subUnitId,
      sheetViewStartRow: row,
      offsetY: rowOffset,
      ...horizontal,
    })
  }
}

/**
 * Minimal-scroll decision for revealing a column on an RTL sheet, in scene
 * space: a column hidden toward the visual left (higher index) becomes flush
 * with the pane's left edge, one hidden toward the visual right (lower
 * index) flush with its right edge. Returns the target viewportScrollX, or
 * null when the column is already fully visible.
 */
export function rtlRevealScrollX(
  cell: { startX: number; endX: number },
  view: { scrollX: number; paneWidth: number },
): number | null {
  const viewLeft = view.scrollX
  const viewRight = view.scrollX + view.paneWidth
  if (cell.startX < viewLeft) return cell.startX
  if (cell.endX > viewRight) return cell.endX - view.paneWidth
  return null
}

/**
 * Reveal-scroll (arrow keys, Ctrl+G, find hits, ScrollToCell). The stock
 * `_scrollToCell` reasons in LTR: "column ≤ first visible → put it at the
 * viewport start" and "column ≥ last visible → walk columns until it fits at
 * the end". On an RTL sheet the first visible column is the visual RIGHT
 * edge, so a column hidden on the right was parked at the far left, and a
 * column hidden on the left (higher index) was walked to a start column that
 * left it still hidden. Keep the stock vertical pass verbatim; decide the
 * horizontal move on mirrored scene edges.
 */
function patchRevealScroll(): void {
  const proto = SheetsScrollRenderController.prototype as any
  const origScrollToCell = proto._scrollToCell
  proto._scrollToCell = function (
    row: number,
    column: number,
    forceTop = false,
    forceLeft = false,
  ) {
    const rtl = rtlMainViewport(
      this._context?.scene,
      this._sheetSkeletonManagerService?.getCurrentSkeleton?.(),
    )
    if (!rtl) return origScrollToCell.call(this, row, column, forceTop, forceLeft)
    const { skeleton, viewMain } = rtl
    const { rowHeightAccumulation, columnWidthAccumulation } = skeleton
    const worksheet = this._context.unit.getActiveSheet()
    if (!rowHeightAccumulation || !columnWidthAccumulation || !worksheet) return false
    row = Math.max(0, Math.min(row, rowHeightAccumulation.length - 1))
    column = Math.max(0, Math.min(column, columnWidthAccumulation.length - 1))
    const {
      startColumn: scrollableStartCol,
      startRow: scrollableStartRow,
      ySplit: freezedRowCount,
    } = worksheet.getFreeze()
    const bounds = this._getViewportBounding()
    if (!bounds) return false
    const { startRow: viewMainStartRow, endRow: viewMainEndRow } = bounds

    // --- vertical: stock logic ---
    let startSheetViewRow: number | undefined
    if (row >= scrollableStartRow && column >= scrollableStartCol - freezedRowCount) {
      if (row <= viewMainStartRow) {
        startSheetViewRow = row
        forceTop = true
      }
      if (row >= viewMainEndRow) {
        const minRowAccumulation = rowHeightAccumulation[row] - viewMain.height
        for (let r = viewMainStartRow; r <= row; r += 1) {
          startSheetViewRow = r + 1
          if (rowHeightAccumulation[r] >= minRowAccumulation) break
        }
      }
    }

    // --- horizontal: mirrored scene edges (forceLeft has no RTL meaning; a
    // visible column stays put) ---
    let horizontal: { sheetViewStartColumn: number; offsetX: number } | undefined
    if (column >= scrollableStartCol && row >= scrollableStartRow - freezedRowCount) {
      const cell = skeleton.getNoMergeCellWithCoordByIndex(row, column)
      const targetX =
        cell && Number.isFinite(cell.startX) && Number.isFinite(cell.endX)
          ? rtlRevealScrollX(cell, { scrollX: viewMain.viewportScrollX, paneWidth: rtl.paneWidth })
          : null
      if (targetX !== null) {
        horizontal = rtlHorizontalScrollState(targetX, rtl, (scrollX) =>
          skeleton.getOffsetRelativeToRowCol(scrollX, 0),
        )
      }
    }
    if (startSheetViewRow === undefined && horizontal === undefined) return false

    const current: {
      offsetY?: number
      sheetViewStartRow?: number
      sheetViewStartColumn?: number
      offsetX?: number
    } = this._scrollManagerService.getCurrentScrollState() || {}
    let offsetY = current.offsetY
    let rowStart: number = startSheetViewRow
      ? Math.min(startSheetViewRow, row)
      : (current.sheetViewStartRow ?? 0) + freezedRowCount
    if (forceTop) {
      offsetY = 0
      rowStart = row
      const hiddenRows = skeleton.getHiddenRowsInRange({
        startRow: rowStart - freezedRowCount,
        endRow: rowStart,
      })
      rowStart -= hiddenRows.length
    }
    return this._commandService.syncExecuteCommand(ScrollCommand.id, {
      sheetViewStartRow: Math.max(0, rowStart - freezedRowCount),
      offsetY,
      ...(horizontal ?? {
        sheetViewStartColumn: current.sheetViewStartColumn,
        offsetX: current.offsetX,
      }),
    })
  }
}

/**
 * Number of frozen columns after dropping the column divider at `column`.
 * Univer derives it from the scroll state's start column (the first visible
 * scrollable column minus the old split); on an RTL sheet that state column
 * is the visually-left (highest) index, so use the LOWEST visible column of
 * the scrollable pane as the anchor instead. Drops inside the frozen band
 * anchor at the first frozen column, as in LTR.
 */
export function rtlFreezeXSplit(params: {
  column: number
  lowestVisibleColumn: number
  oldFreeze: { startColumn: number; xSplit: number }
  droppedInFrozenBand: boolean
}): number {
  const { column, lowestVisibleColumn, oldFreeze, droppedInFrozenBand } = params
  const anchor = droppedInFrozenBand
    ? oldFreeze.startColumn - oldFreeze.xSplit
    : lowestVisibleColumn - oldFreeze.xSplit
  return Math.max(0, column - anchor)
}

const FREEZE_BAR_SIZE = 2
const FREEZE_PERMISSION_KEY = createInterceptorKey('freezePermissionCheck')
const FROZEN_SIDE_VIEWPORTS: string[] = [
  SHEET_VIEWPORT_KEY.VIEW_LEFT_TOP,
  SHEET_VIEWPORT_KEY.VIEW_MAIN_LEFT_TOP,
  SHEET_VIEWPORT_KEY.VIEW_MAIN_LEFT,
  SHEET_VIEWPORT_KEY.VIEW_COLUMN_LEFT,
]

/**
 * Dragging the column freeze divider. The stock pointer handlers snap the
 * preview bar to the hovered column's LTR left edge, clamp it against an LTR
 * "last visible column" and derive the split from the LTR scroll state. Let
 * the stock `_freezeDown` do its permission/cursor setup, then swap its
 * move/up subscriptions for mirrored ones: the bar snaps to the hovered
 * column's mirrored END edge (the frozen band sits to its right) and the
 * split anchors at the lowest visible column. Row-divider drags are untouched.
 */
function patchFreezeDrag(): void {
  const proto = HeaderFreezeRenderController.prototype as any
  const origDown = proto._freezeDown
  proto._freezeDown = function (
    evt: unknown,
    headerRect: any,
    mainRect: any,
    freezeDirectionType = 0,
  ) {
    const skeleton = this._sheetSkeletonManagerService?.getCurrentParam?.()?.skeleton
    const scene = this._getSheetObject?.()?.scene
    const rtl =
      freezeDirectionType === 1 && !!scene && rtlScenes.has(scene) && isRtlSkeleton(skeleton)
    const before = this._scenePointerMoveSub
    const result = origDown.call(this, evt, headerRect, mainRect, freezeDirectionType)
    // Stock bailed out (permission / no skeleton) without subscribing.
    if (!rtl || !this._scenePointerMoveSub || this._scenePointerMoveSub === before) return result
    this._clearObserverEvent()
    const barSize = FREEZE_BAR_SIZE / Math.max(scene.scaleX, scene.scaleY)
    const columnCount = skeleton.columnWidthAccumulation.length
    this._scenePointerMoveSub = scene.onPointerMove$.subscribeEvent((moveEvt: any) => {
      if (!this.interceptor.fetchThroughInterceptors(FREEZE_PERMISSION_KEY)(true, null))
        return false
      const activeViewport = this._getActiveViewport(moveEvt)
      const { row, column: hovered } = getCoordByOffset(
        moveEvt.offsetX,
        moveEvt.offsetY,
        scene,
        skeleton,
        activeViewport || undefined,
        true,
      )
      scene.setCursor(CURSOR_TYPE.GRABBING)
      const column = Math.max(0, Math.min(hovered, columnCount - 1))
      // Mirrored coords: the divider between column-1 (right) and column
      // (left) is the hovered column's mirrored end edge.
      const boundary = this._getPositionByIndex(Math.max(0, row), column)?.endX
      if (typeof boundary === 'number' && Number.isFinite(boundary)) {
        headerRect
          .transformByState({ left: boundary - barSize / 2 })
          ?.setProps({ fill: this._freezeActiveColor })
        mainRect
          .transformByState({ left: boundary - barSize / 2 })
          ?.setProps({ fill: this._freezeNormalHeaderColor })
        this._changeToOffsetX = boundary
      }
      this._changeToColumn = column
      this._activeViewport = activeViewport
    })
    this._scenePointerUpSub = scene.onPointerUp$.subscribeEvent(() => {
      scene.resetCursor()
      scene.enableObjectsEvent()
      this._clearObserverEvent()
      const cancelled = this._changeToColumn === 0 || this._changeToColumn === -1
      headerRect?.setProps({ fill: this._freezeNormalHeaderColor })
      mainRect?.setProps({
        fill: cancelled ? this._freezeNormalMainColor : this._freezeNormalHeaderColor,
      })
      if (cancelled) {
        // Stock ghost-bar spot (kept out of sight on RTL sheets, see patchFreezeBars).
        const ghostLeft = skeleton.rowHeaderWidthAndMarginLeft - barSize
        headerRect?.transformByState({ left: ghostLeft })
        mainRect?.transformByState({ left: ghostLeft })
      }
      const workbook = this._context.unit
      const worksheet = workbook.getActiveSheet()
      const oldFreeze = worksheet?.getConfig?.()?.freeze
      if (!worksheet || !oldFreeze) return
      const viewportKey = this._activeViewport?.viewportKey
      const visible = skeleton.getRangeByViewBound(
        scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)?.calcViewportInfo?.().viewBound,
      )
      const xSplit = rtlFreezeXSplit({
        column: this._changeToColumn,
        lowestVisibleColumn: visible?.startColumn ?? 0,
        oldFreeze,
        droppedInFrozenBand: !viewportKey || FROZEN_SIDE_VIEWPORTS.includes(viewportKey),
      })
      const ySplit = oldFreeze.ySplit || 0
      this._commandService.executeCommand(SetFrozenCommand.id, {
        startRow: ySplit === 0 ? -1 : this._changeToRow,
        startColumn: xSplit === 0 ? -1 : this._changeToColumn,
        ySplit,
        xSplit,
        unitId: workbook.getUnitId(),
        subUnitId: worksheet.getSheetId(),
      })
    })
    return result
  }
}

function patchSkeletonGeometry(): void {
  const proto = SpreadsheetSkeleton.prototype as any

  const origCellByIndex = proto.getCellWithCoordByIndex
  proto.getCellWithCoordByIndex = function (row: number, column: number, header = true) {
    const cell = origCellByIndex.call(this, row, column, header)
    if (!isRtlSkeleton(this) || !cell) return cell
    const offset = header === false ? 0 : this.rowHeaderWidthAndMarginLeft
    return mirrorCellCoord(cell, this.columnTotalWidth, offset)
  }

  const origNoMerge = proto.getNoMergeCellWithCoordByIndex
  proto.getNoMergeCellWithCoordByIndex = function (row: number, column: number, header = true) {
    const cell = origNoMerge.call(this, row, column, header)
    if (!isRtlSkeleton(this) || !cell) return cell
    const offset = header === false ? 0 : this.rowHeaderWidthAndMarginLeft
    return mirrorCellCoord(cell, this.columnTotalWidth, offset)
  }

  // Pointer x → column. The original maps the event offset to grid space via
  // gridX = offsetX / scaleX + scrollX - rowHeaderWidth; feed it a synthetic
  // offset whose grid position is the mirror of the real one.
  const origColByOffset = proto.getColumnIndexByOffsetX
  proto.getColumnIndexByOffsetX = function (
    evtOffsetX: number,
    scaleX: number,
    scrollXY: { x: number; y: number },
    options?: unknown,
  ) {
    if (!isRtlSkeleton(this))
      return origColByOffset.call(this, evtOffsetX, scaleX, scrollXY, options)
    const gridX = evtOffsetX / scaleX + scrollXY.x - this.rowHeaderWidthAndMarginLeft
    const mirroredEvt =
      (this.columnTotalWidth - gridX + this.rowHeaderWidthAndMarginLeft - scrollXY.x) * scaleX
    return origColByOffset.call(this, mirroredEvt, scaleX, scrollXY, options)
  }

  // Scroll-state round trip: viewportScrollX → (column, columnOffset) must
  // invert the mirrored getCellWithCoordByIndex so restores land back on the
  // same pixel.
  const origOffsetRelative = proto.getOffsetRelativeToRowCol
  proto.getOffsetRelativeToRowCol = function (offsetX: number, offsetY: number) {
    if (!isRtlSkeleton(this)) return origOffsetRelative.call(this, offsetX, offsetY)
    const total = this.columnTotalWidth
    const result = origOffsetRelative.call(this, total - offsetX, offsetY)
    const accumulation = this.columnWidthAccumulation
    const mirroredStart = total - (accumulation[result.column] ?? total)
    return { ...result, columnOffset: Math.max(0, offsetX - mirroredStart) }
  }

  // Visible-range culling (and lazy loading through getVisibleRange): the
  // horizontal view window shows the MIRRORED positions of logical columns.
  const origRangeByBounding = proto._getRangeByViewBounding
  proto._getRangeByViewBounding = function (
    rowAcc: number[],
    colAcc: number[],
    viewBound: ViewBoundLike | undefined,
    isPrinting?: boolean,
  ) {
    if (!isRtlSkeleton(this) || !viewBound) {
      return origRangeByBounding.call(this, rowAcc, colAcc, viewBound, isPrinting)
    }
    const offset = this.rowHeaderWidthAndMarginLeft
    const span = mirrorSpanX(viewBound.left, viewBound.right, this.columnTotalWidth, offset)
    return origRangeByBounding.call(
      this,
      rowAcc,
      colAcc,
      { ...viewBound, left: span.startX, right: span.endX },
      isPrinting,
    )
  }
}

function patchSpreadsheetGridlines(): void {
  const proto = Spreadsheet.prototype as any
  const origAuxiliary = proto._drawAuxiliary
  proto._drawAuxiliary = function (ctx: RenderCtxLike) {
    const skeleton = this.getSkeleton()
    if (!isRtlSkeleton(skeleton)) return origAuxiliary.call(this, ctx)
    // Gridlines and merge/overflow clears are text-free: a plain flip lands
    // every line at its mirrored position.
    ctx.save()
    ctx.translate(skeleton.columnTotalWidth, 0)
    ctx.scale(-1, 1)
    origAuxiliary.call(this, ctx)
    ctx.restore()
  }
}

interface BorderCacheValue {
  type: string
  style: number
  color?: string
}

function patchBorderExtension(): void {
  const proto = Border.prototype as any
  const origRender = proto.renderBorderByCell
  proto.renderBorderByCell = function (
    renderBorderContext: { spreadsheetSkeleton: SkeletonLike },
    row: number,
    col: number,
    borderCacheItem: Record<string, BorderCacheValue>,
  ) {
    if (!isRtlSkeleton(renderBorderContext.spreadsheetSkeleton)) {
      return origRender.call(this, renderBorderContext, row, col, borderCacheItem)
    }
    const swapped: Record<string, BorderCacheValue> = {}
    for (const key of Object.keys(borderCacheItem)) {
      const item = borderCacheItem[key]
      if (!item) continue
      const mapped = RTL_BORDER_TYPE_SWAP[item.type]
      swapped[key] = mapped ? { ...item, type: mapped } : item
    }
    return origRender.call(this, renderBorderContext, row, col, swapped)
  }
}

/** Skeleton of the in-flight RTL Font.draw, for helpers without a skeleton arg. */
let activeRtlFontSkeleton: SkeletonLike | null = null

function patchFontExtension(): void {
  const proto = Font.prototype as any

  const origDraw = proto.draw
  proto.draw = function (...args: unknown[]) {
    const skeleton = args[2] as SkeletonLike
    activeRtlFontSkeleton = isRtlSkeleton(skeleton) ? skeleton : null
    try {
      return origDraw.apply(this, args)
    } finally {
      activeRtlFontSkeleton = null
    }
  }

  const origClip = proto._clipRectangleForOverflow
  proto._clipRectangleForOverflow = function (
    ctx: RenderCtxLike,
    startRow: number,
    endRow: number,
    startColumn: number,
    endColumn: number,
    scale: number,
    rowAcc: number[],
    colAcc: number[],
    padding = 0,
  ) {
    if (!activeRtlFontSkeleton) {
      return origClip.call(
        this,
        ctx,
        startRow,
        endRow,
        startColumn,
        endColumn,
        scale,
        rowAcc,
        colAcc,
        padding,
      )
    }
    const startY = rowAcc[startRow - 1] || 0
    const endY = rowAcc[endRow] ?? rowAcc[rowAcc.length - 1] ?? 0
    const startX = colAcc[startColumn - 1] || 0
    const endX = colAcc[endColumn] ?? colAcc[colAcc.length - 1] ?? 0
    const span = mirrorSpanX(startX, endX, activeRtlFontSkeleton.columnTotalWidth)
    ctx.rectByPrecision(
      span.startX + padding,
      startY + padding,
      span.endX - span.startX - 2 * padding,
      endY - startY - 2 * padding,
    )
    ctx.clip()
  }

  // Stock code takes the overflow box from startCell.startX / endCell.endX;
  // mirrored coords reverse that order and produce a negative width, so
  // normalize. Everything else matches the original implementation.
  const origRenderDocuments = proto._renderDocuments
  proto._renderDocuments = function (
    ctx: unknown,
    row: number,
    col: number,
    renderFontCtx: any,
    overflowCache: { getValue(row: number, col: number): Nullable<IRange> },
  ) {
    if (!activeRtlFontSkeleton) {
      return origRenderDocuments.call(this, ctx, row, col, renderFontCtx, overflowCache)
    }
    const documents = this.getDocuments()
    if (documents == null) throw new Error('documents is null')
    const { fontCache } = renderFontCtx
    if (!fontCache) return
    const { documentSkeleton, vertexAngle = 0, wrapStrategy } = fontCache
    if (!documentSkeleton) return
    const documentDataModel = documentSkeleton.getViewModel().getDataModel()
    let { startX, startY, endX, endY } = renderFontCtx
    const cellWidth = endX - startX
    const cellHeight = endY - startY
    if (wrapStrategy === WrapStrategy.WRAP && vertexAngle === 0) {
      documentDataModel.updateDocumentDataPageSize(endX - startX)
      documentSkeleton.calculate()
    } else {
      documentDataModel.updateDocumentDataPageSize(Number.POSITIVE_INFINITY)
    }
    const overflowRectangle = overflowCache.getValue(row, col)
    if (!(wrapStrategy === WrapStrategy.WRAP && vertexAngle === 0) && overflowRectangle) {
      const contentSize = getDocsSkeletonPageSize(documentSkeleton)
      const documentStyle = documentDataModel.getSnapshot().documentStyle
      if (contentSize && documentStyle) {
        const { width } = contentSize
        const { marginRight = 0, marginLeft = 0 } = documentStyle
        documentSkeleton
          .getViewModel()
          .getDataModel()
          .updateDocumentDataPageSize(width + marginLeft + marginRight)
        documentSkeleton.calculate()
      }
      const skeleton = renderFontCtx.spreadsheetSkeleton
      const endCell = skeleton.getCellWithCoordByIndex(
        overflowRectangle.endRow,
        overflowRectangle.endColumn,
      )
      const startCell = skeleton.getCellWithCoordByIndex(
        overflowRectangle.startRow,
        overflowRectangle.startColumn,
      )
      startX = Math.min(startCell.startX, endCell.startX)
      endX = Math.max(startCell.endX, endCell.endX)
      startY = startCell.startY
      endY = endCell.endY
    }
    documentSkeleton.makeDirty(false)
    documents.resize(cellWidth, cellHeight)
    documents.changeSkeleton(documentSkeleton).render(ctx, {
      viewBound: { left: 0, top: 0, right: endX - startX, bottom: endY - startY },
    })
  }
}

/** Flip paint but keep glyphs upright: re-flip around each label's anchor. */
function withMirroredTextCtx(ctx: any, totalWidth: number, run: () => void): void {
  ctx.save()
  ctx.translate(totalWidth, 0)
  ctx.scale(-1, 1)
  const origFillText = ctx.fillText
  ctx.fillText = function (text: string, x: number, y: number) {
    this.save()
    this.translate(x, y)
    this.scale(-1, 1)
    origFillText.call(this, text, 0, 0)
    this.restore()
  }
  try {
    run()
  } finally {
    delete ctx.fillText
    ctx.restore()
  }
}

function patchColumnHeader(): void {
  const layoutProto = ColumnHeaderLayout.prototype as any
  const origLayoutDraw = layoutProto.draw
  layoutProto.draw = function (ctx: any, parentScale: unknown, skeleton: SkeletonLike) {
    if (!isRtlSkeleton(skeleton)) return origLayoutDraw.call(this, ctx, parentScale, skeleton)
    withMirroredTextCtx(ctx, skeleton.columnTotalWidth, () =>
      origLayoutDraw.call(this, ctx, parentScale, skeleton),
    )
  }

  // In RTL the strips are re-anchored, so stray renders into other viewports
  // (which stock geometry happened to clip away) must be filtered.
  const headerProto = SpreadsheetColumnHeader.prototype as any
  const origHeaderDraw = headerProto.draw
  headerProto.draw = function (ctx: unknown, bounds?: { viewportKey?: string }) {
    if (isRtlSkeleton(this.getSkeleton())) {
      const key = bounds?.viewportKey
      if (
        key !== SHEET_VIEWPORT_KEY.VIEW_COLUMN_RIGHT &&
        key !== SHEET_VIEWPORT_KEY.VIEW_COLUMN_LEFT
      )
        return
    }
    return origHeaderDraw.call(this, ctx, bounds)
  }
}

function rowHeaderRtlShift(component: any, skeleton: SkeletonLike): number {
  const scene = component.getScene?.()
  const engine = scene?.getEngine?.()
  if (!engine) return 0
  const scaleX = scene.scaleX || 1
  return Math.max(0, engine.width / scaleX - skeleton.rowHeaderWidthAndMarginLeft)
}

function patchRowHeader(): void {
  const proto = SpreadsheetRowHeader.prototype as any
  const origDraw = proto.draw
  proto.draw = function (ctx: any, bounds?: { viewportKey?: string }) {
    const skeleton = this.getSkeleton()
    if (!isRtlSkeleton(skeleton)) return origDraw.call(this, ctx, bounds)
    const key = bounds?.viewportKey
    if (key !== SHEET_VIEWPORT_KEY.VIEW_ROW_BOTTOM && key !== SHEET_VIEWPORT_KEY.VIEW_ROW_TOP)
      return
    ctx.save()
    ctx.translate(rowHeaderRtlShift(this, skeleton), 0)
    origDraw.call(this, ctx, bounds)
    ctx.restore()
  }

  const origIsHit = proto.isHit
  proto.isHit = function (coord: { x: number; y: number }) {
    const skeleton = this.getSkeleton()
    if (!isRtlSkeleton(skeleton)) return origIsHit.call(this, coord)
    const shift = rowHeaderRtlShift(this, skeleton)
    return origIsHit.call(this, { ...coord, x: coord.x - shift })
  }
}

/**
 * Re-anchor the viewports after every stock reset (`_updateViewport` runs on
 * each skeleton change / freeze change and rewrites all strip rects): main
 * grid gets the left band, the row-header strip and corner move to the right.
 */
function patchFreezeViewportAnchors(): void {
  const proto = HeaderFreezeRenderController.prototype as any
  const origUpdate = proto._updateViewport
  proto._updateViewport = function (
    row = -1,
    column = -1,
    ySplit = 0,
    xSplit = 0,
    resetScroll = 3,
  ) {
    const skeleton = this._sheetSkeletonManagerService?.getCurrentParam?.()?.skeleton
    const scene = this._context?.scene
    const engine = scene?.getEngine?.()
    const rtl = isRtlSkeleton(skeleton)
    // The stock pass must see LTR geometry on an RTL sheet: its column-freeze
    // gap and padding come from getNoMergeCellWithCoordByIndex, and feeding
    // it mirrored scene-space X yields a negative gap and a padding window
    // pinned onto column A's mirrored slice.
    let freezeStartX = 0
    let freezeEndX = 0
    rtlMirrorSuspended = rtl
    try {
      origUpdate.call(this, row, column, ySplit, xSplit, resetScroll)
      if (rtl && column > 0 && skeleton) {
        freezeStartX =
          skeleton.getNoMergeCellWithCoordByIndex(row - ySplit, column - xSplit, false)?.startX ?? 0
        freezeEndX = skeleton.getNoMergeCellWithCoordByIndex(row, column, false)?.startX ?? 0
      }
    } finally {
      rtlMirrorSuspended = false
    }
    if (!scene || !engine) return
    if (!rtl) {
      if (rtlScenes.has(scene)) {
        rtlScenes.delete(scene)
        const viewMain = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
        viewMain?.setMargin(
          skeleton?.rowHeaderWidthAndMarginLeft ?? 0,
          skeleton?.columnHeaderHeightAndMarginTop ?? 0,
        )
        const corner = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_LEFT_TOP)
        if (corner) corner.viewportScrollX = 0
        // A right-edge anchor applied while the RTL flag was mid-switch can
        // leave the LTR sheet scrolled to its far end — replay the stock
        // restore from the recorded scroll state.
        const sm = this._scrollManagerService
        const restored = sm?.calcViewportScrollFromRowColOffset?.(sm?.getCurrentScrollState?.())
        if (viewMain && restored) viewMain.scrollToViewportPos(restored)
        scene.makeDirty(true)
      }
      return
    }
    rtlScenes.add(scene)
    if (this._scrollManagerService) sceneScrollManagers.set(scene, this._scrollManagerService)
    const scaleX = scene.scaleX || 1
    const headerWidth = skeleton.rowHeaderWidthAndMarginLeft
    const { stripLeft, freezeGap, paneLeft, mainRightInset, frozenScrollX } = rtlFreezeLayout({
      engineWidth: engine.width,
      scaleX,
      headerWidth,
      totalWidth: skeleton.columnTotalWidth,
      freezeStartX,
      freezeEndX,
    })
    const viewMain = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
    const viewColumnRight = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_COLUMN_RIGHT)
    const viewRowBottom = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_ROW_BOTTOM)
    const viewRowTop = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_ROW_TOP)
    const viewLeftTop = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_LEFT_TOP)
    viewMain?.setMargin(0, skeleton.columnHeaderHeightAndMarginTop)
    viewMain?.resizeWhenFreezeChange({ left: 0, right: mainRightInset })
    // The stock pass set the pane's horizontal padding to the frozen band's
    // LTR span [bandStart, bandEnd] (header-less x): its scrollbar maps that
    // as "viewportScrollX ≥ bandStart" — the columns hidden behind the band
    // stay unreachable, the band itself is excluded from the range. In the
    // mirrored scene the band and those hidden columns sit at the FAR end,
    // so the equivalent exclusion is a range starting at 0 and shortened by
    // bandEnd (hidden + band); otherwise a freeze dropped while scrolled
    // clamps wheel/reveal onto a mid-grid X and home shows the hidden columns.
    viewMain?.setPadding({
      startX: 0,
      endX: freezeEndX,
      startY: viewMain._paddingStartY,
      endY: viewMain._paddingEndY,
    })
    viewColumnRight?.resizeWhenFreezeChange({ left: 0, right: mainRightInset })
    const viewMainTop = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN_TOP)
    if (viewMainTop?.isActive)
      viewMainTop.resizeWhenFreezeChange({ left: 0, right: mainRightInset })
    if (freezeGap > 0) {
      // Frozen columns dock against the right-edge header strip; their
      // mirrored span sits at the far end of the scene, so the pane scroll
      // pins that span onto the dock.
      for (const key of [
        SHEET_VIEWPORT_KEY.VIEW_MAIN_LEFT,
        SHEET_VIEWPORT_KEY.VIEW_MAIN_LEFT_TOP,
        SHEET_VIEWPORT_KEY.VIEW_COLUMN_LEFT,
      ]) {
        const pane = scene.getViewport(key)
        if (!pane?.isActive) continue
        pane.resizeWhenFreezeChange({ left: paneLeft, width: freezeGap })
        pane.updateScrollVal({ viewportScrollX: frozenScrollX })
      }
    }
    viewRowBottom?.resizeWhenFreezeChange({ left: stripLeft, width: headerWidth + 1 })
    if (viewRowTop?.isActive) viewRowTop.resizeWhenFreezeChange({ left: stripLeft })
    if (viewLeftTop) {
      viewLeftTop.resizeWhenFreezeChange({ left: stripLeft, width: headerWidth + 1 })
      // The corner placeholder Rect stays at scene (-1,-1); shift the corner
      // viewport's world so the Rect lands inside the moved clip.
      viewLeftTop.viewportScrollX = -1 - stripLeft
    }
    filterCornerPlaceholder(this)
    // Load-time restores can run BEFORE these anchors exist, get clamped by
    // stale geometry, and record that clamp as user scroll state. On the
    // first RTL activation of a sheet, wipe the poisoned X-state back to
    // home; then, while at home, anchor flush-right (the patched
    // scrollToViewportPos caps at exactly flush). The stock column reset
    // after a freeze change (resetScroll & 1) is poisoned the same way: it
    // ran under the LTR padding window, whose scrollbar clamped it onto the
    // band's data-space start and recorded that mid-grid column.
    const sheetId = skeleton.worksheet?.getSheetId?.()
    let anchored = anchoredRtlSheets.get(this)
    if (!anchored) anchoredRtlSheets.set(this, (anchored = new Set()))
    const firstActivation = typeof sheetId === 'string' && !anchored.has(sheetId)
    if (firstActivation) anchored.add(sheetId)
    const columnReset = (resetScroll & 1) !== 0
    if (viewMain && (firstActivation || columnReset) && !isAtHorizontalHome(scene)) {
      const state = this._scrollManagerService?.getCurrentScrollState?.() ?? {}
      this._scrollManagerService?.setValidScrollStateToCurrSheet?.({
        ...state,
        sheetViewStartColumn: 0,
        offsetX: 0,
        viewportScrollX: 0,
        viewportScrollY: viewMain.viewportScrollY ?? 0,
      })
    }
    if (viewMain && isAtHorizontalHome(scene)) {
      viewMain.scrollToViewportPos({
        viewportScrollX: Number.MAX_SAFE_INTEGER,
        viewportScrollY: viewMain.viewportScrollY,
      })
    }
    scene.makeDirty(true)
  }
}

/**
 * Layer objects render into every viewport; stock geometry keeps strays out
 * of sight, but the RTL anchors do not, so the corner Rect must render only
 * in its own viewport while an RTL sheet is active.
 */
function filterCornerPlaceholder(controller: any): void {
  const placeholder = controller._context?.components?.get('__SpreadsheetLeftTopPlaceholder__')
  if (!placeholder || placeholder.__rtlViewportFiltered) return
  placeholder.__rtlViewportFiltered = true
  const origRender = placeholder.render.bind(placeholder)
  placeholder.render = (ctx: unknown, vpInfo?: { viewportKey?: string }) => {
    const skeleton = controller._sheetSkeletonManagerService?.getCurrentParam?.()?.skeleton
    if (isRtlSkeleton(skeleton) && vpInfo?.viewportKey !== SHEET_VIEWPORT_KEY.VIEW_LEFT_TOP) {
      return placeholder
    }
    return origRender(ctx, vpInfo)
  }
}
