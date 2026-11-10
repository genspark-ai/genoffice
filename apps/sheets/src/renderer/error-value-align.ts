/**
 * Excel centers error values (#REF!, #N/A, ...) under General alignment,
 * the way it centers booleans; Univer treats them as plain text and
 * left-aligns. A Style-effect interceptor keeps the fix display-only: the
 * stored xf is untouched, so saves and edits never inherit the centering.
 */
import {
  CellValueType,
  HorizontalAlign,
  type ICellData,
  InterceptorEffectEnum,
  type IStyleData,
  type Nullable,
} from '@univerjs/core'
import { ERROR_TYPE_SET, type ErrorType } from '@univerjs/engine-formula'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import type { UniverRuntime } from './univer-state'

export function centeredErrorStyle(
  cell: ICellData,
  style: Nullable<IStyleData>,
): IStyleData | null {
  // A hashed error (#### fill) is already retyped as a number.
  if (cell.p != null || cell.t === CellValueType.NUMBER) return null
  if (typeof cell.v !== 'string' || !ERROR_TYPE_SET.has(cell.v as ErrorType)) return null
  if (style?.ht !== undefined && style.ht !== HorizontalAlign.UNSPECIFIED) return null
  if (style?.tr?.a) return null
  return { ...style, ht: HorizontalAlign.CENTER }
}

export function installErrorValueAlignment(runtime: UniverRuntime): { dispose(): void } {
  const interceptorService = runtime.univer.__getInjector().get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Below the number-format value fixes (9.5) so a hashed error is seen
    // in its final numeric form.
    priority: 9.3,
    effect: InterceptorEffectEnum.Style,
    handler: (cell, location, next) => {
      if (!cell) return next(cell)
      const centered = centeredErrorStyle(cell, location.workbook.getStyles().getStyleByCell(cell))
      return centered === null ? next(cell) : next({ ...cell, s: centered })
    },
  })
}
