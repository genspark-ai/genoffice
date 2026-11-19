/**
 * Excel's Calculation Options. Univer has no runtime auto/manual switch —
 * `initialFormulaComputing` only applies at workbook construction — so Manual
 * is a veto: a beforeCommandExecuted hook cancels the engine's calculation
 * kick-off mutation. (Cancelling — rather than letting the mutation run with
 * an emptied payload — matters: even an empty calculation cycle refreshes
 * volatile functions and fires calculationEnd, which manual mode must not
 * do.) Recalc requests we issue ourselves either carry `forceCalculation`
 * (Calculate Now) or set `allowNext` (Calculate Sheet).
 */
import { ICommandService } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'
import { vetoUniverCommand } from './very-hidden-sheets'

const START_MUTATION = 'formula.mutation.set-formula-calculation-start'

interface CalcState {
  manual: boolean
  allowNext: boolean
}

/// Keyed by runtime so a disposed Univer instance takes its veto hook and
/// mode with it instead of poisoning the next one.
const states = new WeakMap<UniverRuntime, CalcState>()

function stateFor(runtime: UniverRuntime): CalcState {
  const existing = states.get(runtime)
  if (existing) return existing
  const state: CalcState = { manual: false, allowNext: false }
  states.set(runtime, state)
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  commandService.beforeCommandExecuted((command) => {
    if (command.id !== START_MUTATION || !state.manual || state.allowNext) return
    const params = command.params as { forceCalculation?: boolean } | undefined
    if (params?.forceCalculation) return
    vetoUniverCommand(commandService, command, 'manual calculation mode')
  })
  return state
}

export function isManualCalculation(runtime: UniverRuntime | null): boolean {
  return runtime !== null && stateFor(runtime).manual
}

export function setManualCalculation(runtime: UniverRuntime, manual: boolean): void {
  stateFor(runtime).manual = manual
}

/// Opening another file in the same runtime starts back at automatic —
/// calculation mode is workbook state here, not editor state.
export function resetCalculationMode(runtime: UniverRuntime | null): void {
  if (runtime) stateFor(runtime).manual = false
}

export function calculateNow(runtime: UniverRuntime): void {
  runtime.univerAPI.getFormula().executeCalculation()
}

/// F9's little sibling: recalc only the active sheet. dirtyNameMap marks
/// every formula on the sheet dirty (Shift+F9 semantics — formulas whose
/// precedents live on other sheets recalc too); dirtyRanges additionally
/// dirties dependents of this sheet's cells elsewhere in the workbook.
export function calculateSheet(runtime: UniverRuntime): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return
  const sheet = worksheet.getSheet()
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  const unitId = workbook.getId()
  const sheetId = worksheet.getSheetId()
  const state = stateFor(runtime)
  state.allowNext = true
  try {
    commandService.syncExecuteCommand(START_MUTATION, {
      commands: [],
      dirtyRanges: [
        {
          unitId,
          sheetId,
          range: {
            startRow: 0,
            endRow: sheet.getRowCount() - 1,
            startColumn: 0,
            endColumn: sheet.getColumnCount() - 1,
          },
        },
      ],
      dirtyNameMap: { [unitId]: { [sheetId]: worksheet.getSheetName() } },
    })
  } finally {
    state.allowNext = false
  }
}
