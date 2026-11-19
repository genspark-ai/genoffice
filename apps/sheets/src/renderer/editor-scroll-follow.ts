/**
 * Keep the open cell editor on its cell while the grid scrolls.
 *
 * Univer 0.25 places the DOM editor once, when it opens. `EditorContainer`
 * subscribes to `cellEditorManagerService.state$`, but the only writer of that
 * subject is `SheetCellEditorResizeService.fitTextSize()`; the controller's
 * `refreshEditCellPosition()` merely updates the internal layout subject and
 * returns early while the editor is visible. Scrolling therefore moves the
 * grid but leaves the input at its old screen position.
 *
 * Two steps on every scroll frame, in this order:
 *   1. `refreshEditCellPosition()` recomputes the cell rect against the current
 *      viewport scroll and stores it in the internal layout;
 *   2. `fitTextSize()` reads that layout and calls
 *      `cellEditorManagerService.setState()`, which is what moves the DOM.
 *
 * Scroll sources are covered twice over: the viewport's own `onScrollAfter$`
 * catches scrollbar drags (Univer paints its bars on the canvas and drags them
 * with pointer events, so no DOM event fires), keyboard scrolling and any
 * programmatic scroll; document-level `wheel`/`scroll` in the capture phase
 * stay as the fallback for a scroll that happens before the viewport exists.
 */
import { IRenderManagerService } from '@univerjs/engine-render'
import { IEditorBridgeService, SheetCellEditorResizeService } from '@univerjs/sheets-ui'
import type { UniverRuntime } from './univer-state'

export interface EditorScrollFollow {
  /** Re-place the open editor on the next frame (a zoom step just landed). */
  refresh: () => void
  dispose: () => void
}

interface ScrollStream {
  subscribe: (fn: () => void) => { unsubscribe: () => void }
}

export function installEditorScrollFollow(runtime: UniverRuntime): EditorScrollFollow {
  const injector = runtime.univer.__getInjector()
  const bridge = injector.get(IEditorBridgeService)
  const resize = injector.get(SheetCellEditorResizeService)
  if (!bridge || !resize) return { refresh: () => {}, dispose: () => {} }

  const renderManager = injector.get(IRenderManagerService)
  let frame = 0
  let zoomTimer = 0
  let scrollSub: { unsubscribe: () => void } | null = null
  let subscribedStream: ScrollStream | null = null

  function bindViewport(): void {
    if (!renderManager) return
    const unitId = runtime.univerAPI?.getActiveWorkbook?.()?.getId()
    if (!unitId) return
    const render = renderManager.getRenderUnitById(unitId)
    // SHEET_VIEWPORT_KEY.VIEW_MAIN, as the string: the enum is not re-exported
    // from the package root, and the viewport stream is optional at runtime.
    const scene = render?.scene as
      { getViewport?: (key: string) => { onScrollAfter$?: ScrollStream } | null } | null | undefined
    const stream = scene?.getViewport?.('viewMain')?.onScrollAfter$ ?? null
    if (stream === subscribedStream) return
    scrollSub?.unsubscribe()
    scrollSub = stream?.subscribe(() => schedule()) ?? null
    subscribedStream = stream
  }

  function refresh(): void {
    frame = 0
    if (!bridge.isVisible().visible) return
    bridge.refreshEditCellPosition()
    resize.fitTextSize()
  }

  function schedule(): void {
    if (!frame) frame = requestAnimationFrame(refresh)
  }

  bindViewport()

  const onWheel = (event: WheelEvent) => {
    if (event.ctrlKey || event.metaKey) {
      // Ctrl/Command + wheel is the zoom gesture: wheel-zoom.ts steps a notch
      // while no cell is being edited and Univer handles the trackpad pinch.
      // Re-placing the editor on every frame of that gesture fights it, so let
      // it settle and re-place once at the end.
      window.clearTimeout(zoomTimer)
      zoomTimer = window.setTimeout(schedule, 250)
      return
    }
    schedule()
  }
  const onScroll = () => schedule()
  // Opening a file replaces the workbook and its viewport, so the stream bound
  // at install is gone by then; a scrollbar drag starts with a press.
  const onPointerDown = () => bindViewport()

  document.addEventListener('wheel', onWheel, { passive: true, capture: true })
  document.addEventListener('scroll', onScroll, { passive: true, capture: true })
  document.addEventListener('pointerdown', onPointerDown, { passive: true, capture: true })
  return {
    refresh: schedule,
    dispose() {
      if (frame) cancelAnimationFrame(frame)
      window.clearTimeout(zoomTimer)
      scrollSub?.unsubscribe()
      scrollSub = null
      document.removeEventListener('wheel', onWheel, true)
      document.removeEventListener('scroll', onScroll, true)
      document.removeEventListener('pointerdown', onPointerDown, true)
    },
  }
}
