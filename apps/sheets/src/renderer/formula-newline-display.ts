/**
 * A formula whose result carries CR/LF (CHAR(10), a literal typed with a
 * line break) lands in the model as a plain `v`; Univer's plain-text path
 * strips every CR/LF before layout, so the break never renders. File text
 * cells get a rich `p` document at load, but a formula result is rewritten
 * by every recalculation and cannot carry one. Display-only: when the cell
 * wraps, hand the renderer a paragraph document built from the result; with
 * wrap off Univer's stripping already matches Excel's joined single line.
 */
import { type ICellData, type IStyleData, type Nullable, WrapStrategy } from '@univerjs/core'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import { fontTextStyleOf, toRichTextDocument } from './univer-sync'
import type { UniverRuntime } from './univer-state'

/// Shared-formula followers (fill, tiled paste) carry `si` without `f`.
export function isFormulaDriven(cell: Nullable<ICellData>): boolean {
  return (
    (typeof cell?.f === 'string' && cell.f.length > 0) ||
    (typeof cell?.si === 'string' && cell.si.length > 0)
  )
}

export function multilineFormulaDisplayCell(
  cell: ICellData | null | undefined,
  formulaDriven: boolean,
  style: IStyleData | null | undefined,
): ICellData | null | undefined {
  if (!cell || cell.p || !formulaDriven || typeof cell.v !== 'string' || !/[\r\n]/.test(cell.v)) {
    return cell
  }
  if (style?.tb !== WrapStrategy.WRAP) return cell
  return { ...cell, p: toRichTextDocument(cell.v, [], fontTextStyleOf(style)) }
}

export function installFormulaNewlineDisplay(runtime: UniverRuntime): { dispose(): void } {
  const interceptorService = runtime.univer.__getInjector().get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Below NUMFMT (10): the value is final here.
    priority: 9,
    handler: (cell, location, next) => {
      const formulaDriven = isFormulaDriven(cell) || isFormulaDriven(location.rawData)
      if (!formulaDriven || typeof cell?.v !== 'string') return next(cell)
      // getComposedCellStyle goes through getCell and would re-enter this
      // interceptor; the by-cell-data variant composes from the cell in hand.
      const style = location.worksheet.getComposedCellStyleByCellData(
        location.row,
        location.col,
        cell,
      )
      return next(multilineFormulaDisplayCell(cell, formulaDriven, style))
    },
  })
}
