/**
 * What-If Analysis › Goal Seek: solve one input cell so a formula cell hits a
 * target, Excel-style (secant method, Excel's default limits: 100 iterations,
 * |f - target| ≤ 0.001). Each guess is a journaled setValue followed by a
 * recalc wait, so the whole run undoes like typing and works in manual mode.
 */
import { parseAddress } from '../domain/cell-address'
import { calculateNow, isManualCalculation } from './calc-options'
import { t } from './i18n/locale'
import type { UniverRuntime } from './univer-state'

export interface GoalSeekResult {
  readonly found: boolean
  /// f(x) actually reached (shown next to the target).
  readonly reached: number
  readonly solution: number
  readonly iterations: number
}

const MAX_ITERATIONS = 100
const MAX_CHANGE = 0.001

/// Runs `mutate` and resolves after the recalc it triggers. The listener is
/// attached before the write so the end event cannot be missed; the timeout
/// only covers edits that trigger no recalculation at all.
function mutateAndAwaitRecalc(runtime: UniverRuntime, mutate: () => void): Promise<void> {
  const formula = runtime.univerAPI.getFormula()
  return new Promise((resolve) => {
    let settled = false
    let timer: ReturnType<typeof setTimeout> | null = null
    const finish = (): void => {
      if (settled) return
      settled = true
      disposable.dispose()
      if (timer) clearTimeout(timer)
      resolve()
    }
    const disposable = formula.calculationEnd(finish)
    timer = setTimeout(finish, 2_000)
    mutate()
    if (isManualCalculation(runtime)) calculateNow(runtime)
  })
}

export async function solveGoalSeek(
  runtime: UniverRuntime,
  input: { readonly setCell: string; readonly toValue: number; readonly byCell: string },
): Promise<GoalSeekResult> {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!worksheet) throw new Error(t('appNoActiveSheet'))
  const target = parseAddress(input.setCell.toUpperCase())
  const changing = parseAddress(input.byCell.toUpperCase())
  const targetRange = worksheet.getRange(target.row, target.column, 1, 1)
  const changingRange = worksheet.getRange(changing.row, changing.column, 1, 1)
  if (!targetRange.getFormula()) throw new Error(t('appGoalSeekNeedsFormula'))
  // Excel's constraint too: the changing cell must hold a plain number (or
  // nothing) — a formula or text would be destroyed by the solver's writes.
  const originalRaw = changingRange.getValue()
  if (changingRange.getFormula() || (originalRaw != null && typeof originalRaw !== 'number')) {
    throw new Error(t('appGoalSeekByCellNotNumber'))
  }
  const original = typeof originalRaw === 'number' ? originalRaw : 0
  const restore = (): Promise<void> =>
    mutateAndAwaitRecalc(runtime, () =>
      originalRaw == null ? changingRange.setValue({ v: null }) : changingRange.setValue(original),
    )

  /// NaN when the formula errors for this guess (#DIV/0! mid-solve is
  /// normal); the secant guards treat it as a dead end, not a crash.
  const evaluate = async (x: number): Promise<number> => {
    await mutateAndAwaitRecalc(runtime, () => changingRange.setValue(x))
    const value = targetRange.getValue()
    return typeof value === 'number' ? value : Number.NaN
  }

  try {
    let x0 = original
    let f0 = await evaluate(x0)
    if (Number.isNaN(f0)) throw new Error(t('appGoalSeekTargetNotNumber'))
    if (Math.abs(f0 - input.toValue) <= MAX_CHANGE) {
      return { found: true, reached: f0, solution: x0, iterations: 1 }
    }
    // A relative nudge for the second sample; 1 covers the x0 = 0 case.
    let x1 = x0 + Math.max(Math.abs(x0) * 0.01, 1)
    let f1 = await evaluate(x1)
    let iterations = 2
    while (iterations < MAX_ITERATIONS) {
      if (Math.abs(f1 - input.toValue) <= MAX_CHANGE) {
        return { found: true, reached: f1, solution: x1, iterations }
      }
      const slope = f1 - f0
      if (slope === 0 || !Number.isFinite(slope)) break
      const next = x1 - ((f1 - input.toValue) * (x1 - x0)) / slope
      if (!Number.isFinite(next)) break
      x0 = x1
      f0 = f1
      x1 = next
      f1 = await evaluate(x1)
      iterations += 1
    }
    if (Math.abs(f1 - input.toValue) <= MAX_CHANGE) {
      return { found: true, reached: f1, solution: x1, iterations }
    }
    // No convergence: put the original contents back (journaled, undoable).
    await restore()
    return { found: false, reached: f1, solution: x1, iterations }
  } catch (error) {
    // Any abort after the first write must not leave a guess in the grid.
    await restore()
    throw error
  }
}
