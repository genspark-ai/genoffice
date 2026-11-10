/**
 * Cell affordances for threaded comments: Excel paints a purple corner
 * triangle (notes keep the red/yellow one) and shows the thread on hover.
 * The marker rides Univer's cell-content interceptor like the note marker;
 * hover goes through the sheet hover service so the card follows the cell
 * under the pointer, not the DOM.
 */
import { InterceptorEffectEnum } from '@univerjs/core'
import { IRenderManagerService } from '@univerjs/engine-render'
import { SheetSkeletonManagerService } from '@univerjs/preset-sheets-core'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'
import { HoverManagerService } from '@univerjs/sheets-ui'

import { threadStore, type CellThread } from './threaded-comments'
import type { UniverRuntime } from './univer-state'

/** Canvas chrome colors, kept out of document data and keyed by UI theme. */
export const THREAD_MARKER_CANVAS_COLORS = {
  light: { marker: '#7b3fe4' },
  dark: { marker: '#a78bfa' },
} as const

export type ThreadMarkerTheme = keyof typeof THREAD_MARKER_CANVAS_COLORS

export interface ThreadHover {
  readonly sheetId: string
  readonly thread: CellThread
  readonly clientX: number
  readonly clientY: number
}

interface Options {
  readonly theme: () => ThreadMarkerTheme
  readonly onHover: (hover: ThreadHover | null) => void
}

export function installThreadedCommentAffordances(
  runtime: UniverRuntime,
  options: Options,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptor = injector
    .get(SheetInterceptorService)
    .intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
      // Above the note marker (100) so a thread wins the corner.
      priority: 101,
      effect: InterceptorEffectEnum.Style,
      handler: (cell, position, next) => {
        if (!threadStore.get(position.subUnitId, position.row, position.col)) return next(cell)
        const base = !cell || cell === position.rawData ? { ...position.rawData } : cell
        return next({
          ...base,
          markers: {
            ...base?.markers,
            tr: { color: THREAD_MARKER_CANVAS_COLORS[options.theme()].marker, size: 6 },
          },
        })
      },
    })

  const refresh = (): void => {
    const workbookId = runtime.univerAPI.getActiveWorkbook()?.getId()
    if (!workbookId) return
    try {
      const render = injector.get(IRenderManagerService).getRenderById(workbookId)
      render?.with(SheetSkeletonManagerService).reCalculate()
      render?.mainComponent?.makeDirty(true)
      render?.scene?.makeDirty(true)
    } catch {
      // Render modules may not exist yet during workbook replacement.
    }
  }
  const unsubscribeStore = threadStore.subscribe(refresh)

  let lastKey = ''
  const hover = injector.get(HoverManagerService).currentCellPosWithEvent$.subscribe((cell) => {
    if (!cell) {
      if (lastKey) options.onHover(null)
      lastKey = ''
      return
    }
    const thread = threadStore.get(cell.subUnitId, cell.row, cell.col)
    const key = thread ? `${cell.subUnitId}:${cell.row}:${cell.col}` : ''
    if (key === lastKey) return
    lastKey = key
    const event = cell.event as { clientX?: number; clientY?: number }
    options.onHover(
      thread
        ? {
            sheetId: cell.subUnitId,
            thread,
            clientX: event.clientX ?? 0,
            clientY: event.clientY ?? 0,
          }
        : null,
    )
  })

  return {
    dispose() {
      interceptor.dispose()
      unsubscribeStore()
      hover.unsubscribe()
    },
  }
}
