import type { UniverRuntime } from './univer-state'

/// A cold load reaches preloadComplete before the formula engine's first pass
/// lands in the model (the full recalculation the stream hold releases comes
/// a second later), so an export right after it writes formula cells empty.
/// Force one calculation pass, wait for its end (bounded) and then let the
/// model absorb the results (also bounded) — calculationEnd alone was not
/// enough on a cold start.
export const FORMULA_SETTLE_TIMEOUT_MS = 30_000

export async function settleFormulas(runtime: UniverRuntime | null): Promise<void> {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!runtime || !workbook) return
  /** formula cells of the whole workbook whose computed value has not landed */
  const pendingFormulas = (): number => {
    const snapshot = workbook.getSnapshot() as unknown as {
      sheets?: Record<
        string,
        {
          cellData?: Record<string, Record<string, { f?: unknown; v?: unknown }> | undefined>
        }
      >
    } | null
    let pending = 0
    for (const sheet of Object.values(snapshot?.sheets ?? {})) {
      for (const row of Object.values(sheet?.cellData ?? {})) {
        for (const cell of Object.values(row ?? {})) {
          if (cell?.f && (cell.v === undefined || cell.v === null)) pending += 1
        }
      }
    }
    return pending
  }
  if (pendingFormulas() === 0) return
  const formula = runtime.univerAPI.getFormula()
  const ended = new Promise<void>((resolve) => {
    const disposable = formula.calculationEnd(() => {
      disposable?.dispose()
      resolve()
    })
  })
  formula.executeCalculation()
  await Promise.race([
    ended,
    new Promise<void>((resolve) => setTimeout(resolve, FORMULA_SETTLE_TIMEOUT_MS)),
  ])
  const deadline = Date.now() + FORMULA_SETTLE_TIMEOUT_MS
  while (Date.now() < deadline && pendingFormulas() > 0) {
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
}
