/**
 * Excel never flags a cell that fails its data-validation rule on screen; the
 * red "invalid data" mark only appears when the user runs Circle Invalid
 * Data. Univer's render controller paints a red corner triangle on every
 * INVALID cell, so a workbook whose list source we cannot evaluate (defined
 * name, structured reference, #REF!, a range outside the streamed window)
 * shows a false marker on the whole validated column. Strip that marker
 * after Univer adds it; dropdowns, input messages and edit-time rejection
 * are untouched.
 */
import { InterceptorEffectEnum } from '@univerjs/core'
import type { ICellDataForSheetInterceptor, Nullable } from '@univerjs/core'
import {
  INTERCEPTOR_POINT,
  InterceptCellContentPriority,
  SheetInterceptorService,
} from '@univerjs/sheets'

import type { UniverRuntime } from './univer-state'

/// Univer's INVALID_MARK (dv-render.controller.ts) — the only red `tr` marker
/// the sheet paints.
const INVALID_MARK_COLOR = '#fe4b4b'

export function stripInvalidDataMarker(
  cell: Nullable<ICellDataForSheetInterceptor>,
): Nullable<ICellDataForSheetInterceptor> {
  const marker = cell?.markers?.tr
  if (!marker || marker.color.toLowerCase() !== INVALID_MARK_COLOR) return cell
  const { tr: _dropped, ...rest } = cell.markers ?? {}
  return { ...cell, markers: rest }
}

export function installInvalidDataMarkerSuppression(runtime: UniverRuntime): {
  dispose(): void
} {
  return runtime.univer
    .__getInjector()
    .get(SheetInterceptorService)
    .intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
      // Just below the data-validation interceptor so this sees its marker.
      priority: InterceptCellContentPriority.DATA_VALIDATION - 1,
      effect: InterceptorEffectEnum.Style,
      handler: (cell, _position, next) => next(stripInvalidDataMarker(cell)),
    })
}
