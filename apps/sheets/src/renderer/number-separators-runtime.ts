/**
 * Hooks the region's number separators into Univer: display text through the
 * numfmt controller's locale, and the cell editor / formula bar through the
 * edit write chain, so the model only ever holds canonical values and formulas.
 */
import { CellValueType, type ICellData, type Nullable } from '@univerjs/core'
import { AFTER_CELL_EDIT, BEFORE_CELL_EDIT, SheetInterceptorService } from '@univerjs/sheets'
import { SheetsNumfmtCellContentController } from '@univerjs/sheets-numfmt'

import {
  canonicalNumberToLocal,
  formulaToCanonical,
  formulaToLocal,
  getNumberSeparators,
  localNumberToCanonical,
  numfmtDisplayLocale,
  onNumberSeparatorsChange,
  usesDecimalComma,
} from './number-separators'
import type { UniverRuntime } from './univer-state'

type Cell = Nullable<ICellData>

/** The cell as the editor opens it: formula and number text in the region's notation. */
export function cellForEditor(cell: Cell): Cell {
  const sep = getNumberSeparators()
  if (!cell || !usesDecimalComma(sep)) return cell
  if (typeof cell.f === 'string' && cell.f !== '') {
    const formula = formulaToLocal(cell.f, sep)
    return formula === cell.f ? cell : { ...cell, f: formula }
  }
  const numeric = typeof cell.v === 'number' || cell.t === CellValueType.NUMBER
  if (!numeric || cell.p != null || (typeof cell.v !== 'number' && typeof cell.v !== 'string'))
    return cell
  const text = canonicalNumberToLocal(cell.v, sep)
  return text === String(cell.v) ? cell : { ...cell, v: text }
}

/** The committed edit rewritten in canonical notation, before Univer parses it. */
export function cellFromEditor(cell: Cell): Cell {
  const sep = getNumberSeparators()
  if (!cell || !usesDecimalComma(sep)) return cell
  if (typeof cell.f === 'string' && cell.f !== '') {
    const formula = formulaToCanonical(cell.f, sep)
    return formula === cell.f ? cell : { ...cell, f: formula }
  }
  // text-formatted and apostrophe-forced cells keep what was typed
  if (cell.t === CellValueType.STRING || cell.t === CellValueType.FORCE_STRING) return cell
  if (typeof cell.v !== 'string' || cell.p != null) return cell
  const canonical = localNumberToCanonical(cell.v, sep)
  return canonical === null ? cell : { ...cell, v: canonical }
}

export function installNumberSeparators(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptorService = injector.get(SheetInterceptorService)
  const numfmtController = injector.get(SheetsNumfmtCellContentController)
  type LocaleTag = Parameters<typeof numfmtController.setNumfmtLocal>[0]
  const syncLocale = () => numfmtController.setNumfmtLocal(numfmtDisplayLocale() as LocaleTag)
  syncLocale()
  const stopWatch = onNumberSeparatorsChange(syncLocale)
  // Univer's formula hook ends the BEFORE_CELL_EDIT chain with the model's
  // formula, so no interceptor can run after it: wrap the composed chain
  const writeChain = interceptorService.writeCellInterceptor
  const fetchChain = writeChain.fetchThroughInterceptors
  const wrapped = ((name: unknown, filter?: unknown) => {
    const composed = (fetchChain as (...args: unknown[]) => unknown).call(writeChain, name, filter)
    if (name !== BEFORE_CELL_EDIT) return composed
    return (value: Cell, context: unknown) =>
      cellForEditor((composed as (value: Cell, context: unknown) => Cell)(value, context))
  }) as typeof fetchChain
  writeChain.fetchThroughInterceptors = wrapped
  // first in the chain: Univer's number and formula parsing sees canonical text
  const afterEdit = interceptorService.writeCellInterceptor.intercept(AFTER_CELL_EDIT, {
    priority: 100,
    handler: (value, _context, next) => next(cellFromEditor(value)),
  })
  return {
    dispose() {
      stopWatch()
      if (writeChain.fetchThroughInterceptors === wrapped)
        writeChain.fetchThroughInterceptors = fetchChain
      afterEdit()
    },
  }
}
