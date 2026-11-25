import type { Worksheet } from '@univerjs/core'
import { SHEET_VIEWPORT_KEY, Spreadsheet, Viewport } from '@univerjs/engine-render'
import { SHEET_VIEW_KEY } from '@univerjs/sheets-ui'

const LARGE_ROW_COUNT = 100_000

let deferredVerticalThumbInstalled = false

function viewportWorksheet(viewport: Viewport): Worksheet | undefined {
  const grid = viewport.scene.getObject(SHEET_VIEW_KEY.MAIN)
  return grid instanceof Spreadsheet ? grid.getSkeleton()?.worksheet : undefined
}

/** Paint the thumb while dragging on tall sheets and scroll once on release. */
export function installDeferredVerticalThumb(): void {
  if (deferredVerticalThumbInstalled) return
  deferredVerticalThumbInstalled = true

  const prototype = Viewport.prototype
  const scrollByDelta = prototype.scrollByBarDeltaValue
  const dispose = prototype.dispose
  const drags = new WeakMap<Viewport, { targetY: number; finish(commit: boolean): void }>()

  prototype.dispose = function () {
    drags.get(this)?.finish(false)
    dispose.call(this)
  }

  prototype.scrollByBarDeltaValue = function (delta, isTrigger) {
    // Wheel deltas always include x. The vertical thumb owns the pointer
    // while its private flag is set, excluding selection auto-scroll.
    if (
      this.viewportKey !== SHEET_VIEWPORT_KEY.VIEW_MAIN ||
      delta.y === undefined ||
      delta.x !== undefined
    ) {
      return scrollByDelta.call(this, delta, isTrigger)
    }
    const scrollbar = this.getScrollBar()
    const dragState = scrollbar as unknown as { _isVerticalMove: boolean } | null
    if (!scrollbar || !dragState?._isVerticalMove) return scrollByDelta.call(this, delta, isTrigger)

    let drag = drags.get(this)
    if (!drag) {
      const sheet = viewportWorksheet(this)
      if (!sheet || sheet.getRowCount() < LARGE_ROW_COUNT) {
        return scrollByDelta.call(this, delta, isTrigger)
      }
      const listeners = new AbortController()
      const renderThumb = scrollbar.render
      const pending = {
        targetY: this.scrollY,
        finish: (commit: boolean): void => {
          if (!drags.delete(this)) return
          listeners.abort()
          scrollbar.render = renderThumb
          // Release can land outside the canvas or follow a cancellation.
          if (dragState._isVerticalMove) {
            dragState._isVerticalMove = false
            this.scene.releaseCapturedObject()
            this.scene.enableObjectsEvent()
          }
          if (commit && this.isActive && viewportWorksheet(this) === sheet) {
            this.scrollToBarPos({ y: pending.targetY })
          } else {
            scrollbar.makeViewDirty(true)
          }
        },
      }
      drags.set(this, pending)
      drag = pending
      // Only the scrollbar renderer sees the pending coordinate. Keeping
      // viewportScrollY untouched avoids scroll events and style cache work.
      scrollbar.render = (...args) => {
        const originalY = this.scrollY
        this.scrollY = pending.targetY
        try {
          renderThumb.apply(scrollbar, args)
        } finally {
          this.scrollY = originalY
        }
      }
      if (this.scrollAnimationFrameId !== null) {
        cancelAnimationFrame(this.scrollAnimationFrameId)
        this.scrollAnimationFrameId = null
      }
      const options = { capture: true, signal: listeners.signal }
      window.addEventListener(
        'pointerup',
        () => queueMicrotask(() => pending.finish(true)),
        options,
      )
      window.addEventListener('pointercancel', () => pending.finish(false), options)
      window.addEventListener('blur', () => pending.finish(false), options)
    }
    const position = this.limitedScroll(this.scrollX, drag.targetY + delta.y)
    drag.targetY = position.scrollY
    scrollbar.makeViewDirty(true)
    return position
  }
}
