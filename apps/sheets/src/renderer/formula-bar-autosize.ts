/**
 * Formula-bar auto-height — WPS/Excel grow the fx bar when the selected
 * cell's formula wraps; Univer keeps a fixed 28px row and clips mid-line.
 *
 * Univer's FormulaBar owns its height through Tailwind classes
 * (`univer-h-7` collapsed / `univer-h-20` arrow-expanded). This module
 * measures the fx-bar editor doc's real content height from its skeleton
 * and overrides the bar's inline height to fit, capped so a monster
 * formula cannot eat the grid. The workbench is a flex column, so the
 * grid reflows on its own, and the bar's own ResizeObserver re-anchors
 * the editor canvas.
 */
import { DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY, ICommandService } from '@univerjs/core'
import { DocSkeletonManagerService, RichTextEditingMutation } from '@univerjs/docs'
import { IRenderManagerService } from '@univerjs/engine-render'
import { IEditorBridgeService } from '@univerjs/sheets-ui'

import type { UniverRuntime } from './univer-state'

/// Univer's own heights: h-7 collapsed, h-20 after the expand arrow.
const COLLAPSED_PX = 28
const EXPANDED_PX = 80
/// The doc paints from the bar's top edge; this keeps a single line's
/// position identical to the stock 28px bar when a second line appears.
const VERTICAL_PAD_PX = 9
/// Never grow past ~5 wrapped lines; a longer formula clips at a clean
/// line boundary and the expand arrow still offers the 80px mode.
const MAX_PX = 124

export function installFormulaBarAutosize(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const renderManagerService = injector.get(IRenderManagerService)
  const commandService = injector.get(ICommandService)
  const bridge = injector.get(IEditorBridgeService)

  let disposed = false
  let raf = 0
  let settleTimer: ReturnType<typeof setTimeout> | undefined
  let appliedPx: number | null = null
  let observedBar: HTMLElement | null = null
  let barObserver: ResizeObserver | null = null
  let classObserver: MutationObserver | null = null

  const contentHeight = (): number | null => {
    const skeleton = renderManagerService
      .getRenderById(DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY)
      ?.with(DocSkeletonManagerService)
      .getSkeleton()
    const page = skeleton?.getSkeletonData()?.pages[0]
    return typeof page?.height === 'number' && Number.isFinite(page.height) ? page.height : null
  }

  // Re-fit when the arrow toggles h-7/h-20, re-wrap when the bar resizes
  // horizontally. The element can remount with the workbench, so re-bind
  // on every measure.
  const watch = (bar: HTMLElement) => {
    if (observedBar === bar) return
    barObserver?.disconnect()
    classObserver?.disconnect()
    observedBar = bar
    // A fresh node carries no inline height; forget the cache so the next
    // measure writes it again.
    appliedPx = null
    barObserver = new ResizeObserver(schedule)
    barObserver.observe(bar)
    classObserver = new MutationObserver(schedule)
    classObserver.observe(bar, { attributes: true, attributeFilter: ['class'] })
  }

  const measure = () => {
    const bar = document.querySelector<HTMLElement>('[data-u-comp="formula-bar"]')
    if (!bar) return
    watch(bar)
    if (bar.offsetParent === null) return // View > Formula Bar off
    // Growing the bar mid-edit reflows the grid under the cell editor
    // overlay; freeze until the session ends (the arrow still works).
    if (bridge.isVisible().visible) return
    const height = contentHeight()
    if (height === null) return
    const base = bar.classList.contains('univer-h-20') ? EXPANDED_PX : COLLAPSED_PX
    const fit = Math.ceil(height) + VERTICAL_PAD_PX
    const target = Math.min(Math.max(fit, base), MAX_PX)
    const next = target > base ? target : null
    if (next === appliedPx) return
    appliedPx = next
    bar.style.height = next === null ? '' : `${next}px`
  }

  const schedule = () => {
    if (disposed) return
    cancelAnimationFrame(raf)
    // Post-layout: the doc skeleton recalculates in the render loop, so
    // read it a frame after the mutation that triggered us.
    raf = requestAnimationFrame(() => {
      raf = requestAnimationFrame(measure)
    })
  }

  // Selection changes rewrite the fx-bar preview (sometimes a tick later,
  // via the formula-text cover in formula-view.ts) — settle twice.
  const editCellSub = bridge.currentEditCellState$.subscribe(() => {
    schedule()
    clearTimeout(settleTimer)
    settleTimer = setTimeout(schedule, 30)
  })
  // Edit session end unfreezes the height (the cell may have a new formula).
  const visibleSub = bridge.visible$.subscribe((state) => {
    if (!state.visible) schedule()
  })
  // Any direct rewrite of the fx-bar doc (cover command, fx button).
  const commandSub = commandService.onCommandExecuted((command) => {
    if (
      command.id === RichTextEditingMutation.id &&
      (command.params as { unitId?: string } | undefined)?.unitId ===
        DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY
    ) {
      schedule()
    }
  })

  return {
    dispose() {
      disposed = true
      cancelAnimationFrame(raf)
      clearTimeout(settleTimer)
      editCellSub.unsubscribe()
      visibleSub.unsubscribe()
      commandSub.dispose()
      barObserver?.disconnect()
      classObserver?.disconnect()
      if (observedBar) observedBar.style.height = ''
    },
  }
}
