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
 * Column freeze on an RTL sheet is not mirrored yet (rare; row freeze works).
 */
import { BooleanNumber, WrapStrategy } from '@univerjs/core'
import type { IRange, Nullable } from '@univerjs/core'
import {
  Border,
  ColumnHeaderLayout,
  Font,
  getDocsSkeletonPageSize,
  SHEET_VIEWPORT_KEY,
  Spreadsheet,
  SpreadsheetColumnHeader,
  SpreadsheetRowHeader,
  SpreadsheetSkeleton,
  Viewport,
} from '@univerjs/engine-render'
import { HeaderFreezeRenderController } from '@univerjs/sheets-ui'

/** Scenes whose active sheet is RTL — consulted by shared prototype patches. */
const rtlScenes = new WeakSet<object>()

/** Scene → scroll manager, so viewport patches can ask "is this sheet at home?". */
const sceneScrollManagers = new WeakMap<object, any>()

let reapplyingRtlMainScroll = false

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

interface SkeletonLike {
  worksheet?: { getConfig?(): { rightToLeft?: BooleanNumber } } | null
  columnTotalWidth: number
  columnWidthAccumulation: number[]
  rowHeaderWidthAndMarginLeft: number
  columnHeaderHeightAndMarginTop: number
}

export function isRtlSkeleton(skeleton: SkeletonLike | null | undefined): boolean {
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
      const maxUseful = this._scene.width - (this.width || 0) / scaleX
      if (maxUseful > 0 && pos.viewportScrollX > maxUseful) {
        // An overscrolled request is the RTL "home" restore (state A1/0 maps
        // past the edge). Cap it, and keep it out of the recorded scroll
        // state (isTrigger=false) so later restores re-cap against fresh
        // geometry instead of freezing a stale clamp.
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
    origUpdate.call(this, row, column, ySplit, xSplit, resetScroll)
    const skeleton = this._sheetSkeletonManagerService?.getCurrentParam?.()?.skeleton
    const scene = this._context?.scene
    const engine = scene?.getEngine?.()
    if (!scene || !engine) return
    if (!isRtlSkeleton(skeleton)) {
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
    const stripLeft = engine.width / scaleX - headerWidth - 1
    const viewMain = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
    const viewColumnRight = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_COLUMN_RIGHT)
    const viewRowBottom = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_ROW_BOTTOM)
    const viewRowTop = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_ROW_TOP)
    const viewLeftTop = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_LEFT_TOP)
    viewMain?.setMargin(0, skeleton.columnHeaderHeightAndMarginTop)
    viewMain?.resizeWhenFreezeChange({ left: 0, right: (headerWidth + 1) * scaleX })
    viewColumnRight?.resizeWhenFreezeChange({ left: 0, right: (headerWidth + 1) * scaleX })
    const viewMainTop = scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN_TOP)
    if (viewMainTop?.isActive)
      viewMainTop.resizeWhenFreezeChange({ left: 0, right: (headerWidth + 1) * scaleX })
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
    // scrollToViewportPos caps at exactly flush).
    const sheetId = skeleton.worksheet?.getSheetId?.()
    let anchored = anchoredRtlSheets.get(this)
    if (!anchored) anchoredRtlSheets.set(this, (anchored = new Set()))
    const firstActivation = typeof sheetId === 'string' && !anchored.has(sheetId)
    if (firstActivation) anchored.add(sheetId)
    if (viewMain && firstActivation && !isAtHorizontalHome(scene)) {
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
