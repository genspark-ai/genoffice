/**
 * Display fallback to the file's cached formula results.
 *
 * When the engine's recalculation of a file formula errors — a function it
 * lacks or a name it can't resolve — Excel's cached result from the file is
 * the better display. Display-only: the model keeps the engine's result and
 * the chain continues, so number formats still apply. A cell the user
 * overwrote this session shows the live engine truth instead.
 */
import { CellValueType } from '@univerjs/core'
import { ERROR_TYPE_SET, ErrorType } from '@univerjs/engine-formula'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/// Journal ops that renumber rows/columns; the recorded cache coordinates are
/// stale after these. Layout-only ops (sizing, hiding, outlining, merging)
/// leave addresses intact.
const COORDINATE_SHIFTING_OPS = new Set([
  'insert-rows',
  'remove-rows',
  'insert-cols',
  'remove-cols',
  'move-rows',
])

/**
 * #VALUE! from a formula built only of literals, references and operators is
 * the engine's own arithmetic verdict (text the coercion rules reject, like
 * Excel recalculating "76,92"+…), not a capability gap — the file cache must
 * not mask it. Functions, defined names and structured/external refs keep
 * the fallback: those are where the engine errors where Excel would not.
 */
export function isPlainArithmeticFormula(formula: string): boolean {
  let body = formula.startsWith('=') ? formula.slice(1) : formula
  body = body.replace(/"(?:[^"]|"")*"/g, '""').replace(/\$/g, '')
  // Structured/external refs, array literals and spill refs are not "plain".
  if (/[[{#]/.test(body)) return false
  body = body.replace(/'(?:[^']|'')*'!/g, '').replace(/[A-Za-z_][\w.]*!/g, '')
  // Scientific literals (1E-3) would otherwise shed a bare E token.
  body = body.replace(/(?<![\w.])\d+(?:\.\d+)?E[+-]?\d+/gi, '0')
  for (const match of body.matchAll(/[A-Za-z_][\w.]*/g)) {
    const token = match[0]
    // A ref-shaped name followed by "(" is a function call (LOG10, ...).
    if (/^\s*\(/.test(body.slice((match.index ?? 0) + token.length))) return false
    if (!/^[A-Za-z]{1,3}\d+$/.test(token) && !/^(?:TRUE|FALSE)$/i.test(token)) return false
  }
  return true
}

export function installCachedValueFallbackInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const interceptorService = runtime.univer.__getInjector().get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Below the formula-view interceptors (formula view must keep showing
    // the formula text), above NUMFMT (10).
    priority: 9997,
    handler: (cell, location, next) => {
      const value = cell?.v
      if (typeof value !== 'string' || !ERROR_TYPE_SET.has(value as ErrorType)) {
        return next(cell)
      }
      if (value === ErrorType.VALUE) {
        const formula = location.rawData?.f
        if (typeof formula === 'string' && isPlainArithmeticFormula(formula)) return next(cell)
      }
      const state = lazyWorkbookRef.current
      if (!state) return next(cell)
      const sheetId = location.subUnitId
      const ops = state.editJournal.structuralOps.get(sheetId)
      // Row/column shifts invalidate the recorded coordinates; user
      // overwrites own the cell.
      if (ops?.some((op) => COORDINATE_SHIFTING_OPS.has(op.kind))) return next(cell)
      const key = `${location.row}:${location.col}`
      const edits = state.editJournal.cells.get(sheetId)
      // Style-only journal entries (hasValue false) don't change any input.
      if (edits?.get(key)?.hasValue) return next(cell)
      // Once the sheet has user content edits, a computed error (#DIV/0!,
      // #N/A, ...) may be the true result of the new inputs — keep it.
      // #NAME? still falls back: the engine cannot evaluate that formula
      // under any inputs, so the file's value remains the best display.
      if (value !== ErrorType.NAME && edits) {
        for (const entry of edits.values()) {
          if (entry.hasValue) return next(cell)
        }
      }
      const cached = state.cachedFormulaValues.get(sheetId)?.get(key)
      if (cached === undefined || cached === value) return next(cell)
      return next({
        ...cell,
        v: cached,
        t:
          typeof cached === 'number'
            ? CellValueType.NUMBER
            : typeof cached === 'boolean'
              ? CellValueType.BOOLEAN
              : CellValueType.STRING,
      })
    },
  })
}
