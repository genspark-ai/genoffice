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
 * Listening on the document in the capture phase catches wheel scrolling over
 * the canvas, Univer's own scrollbar (scroll does not bubble) and any other
 * scroll source, even when an inner handler stops propagation.
 */
import { IEditorBridgeService, SheetCellEditorResizeService } from '@univerjs/sheets-ui'
import type { UniverRuntime } from './univer-state'

export interface EditorScrollFollow {
  /** Re-place the open editor on the next frame (a zoom step just landed). */
  refresh: () => void
  dispose: () => void
}

export function installEditorScrollFollow(runtime: UniverRuntime): EditorScrollFollow {
  const injector = runtime.univer.__getInjector()
  const bridge = injector.get(IEditorBridgeService)
  const resize = injector.get(SheetCellEditorResizeService)
  if (!bridge || !resize) return { refresh: () => {}, dispose: () => {} }

  let frame = 0
  let zoomTimer = 0
  const refresh = () => {
    frame = 0
    if (!bridge.isVisible().visible) return
    bridge.refreshEditCellPosition()
    resize.fitTextSize()
  }
  const schedule = () => {
    if (!frame) frame = requestAnimationFrame(refresh)
  }
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

  document.addEventListener('wheel', onWheel, { passive: true, capture: true })
  document.addEventListener('scroll', onScroll, { passive: true, capture: true })
  return {
    refresh: schedule,
    dispose() {
      if (frame) cancelAnimationFrame(frame)
      window.clearTimeout(zoomTimer)
      document.removeEventListener('wheel', onWheel, true)
      document.removeEventListener('scroll', onScroll, true)
    },
  }
}
