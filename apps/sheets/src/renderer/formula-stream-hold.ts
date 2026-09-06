/**
 * Univer's TriggerCalculationController turns every set-range-values mutation
 * into a formula calculation cycle after a 100 ms debounce, and a cycle that
 * is still running when the next one arrives is stopped and restarted. Each
 * cycle rebuilds the dependency tree of every formula in the workbook (the
 * AST cache holds 5,000 entries), so a streamed load or a scroll through an
 * unloaded area of a workbook with thousands of formulas spends seconds per
 * chunk in the engine and never finishes a cycle (conditional formats that
 * depend on it stay unpainted). File data carries its cached values and needs
 * no recalculation per chunk: hold the engine's kick-off while chunks are
 * landing and release the accumulated dirty ranges as one cycle afterwards.
 */
import { CustomCommandExecutionError, ICommandService } from '@univerjs/core'
import { TriggerCalculationController } from '@univerjs/sheets-formula'

import type { UniverRuntime } from './univer-state'

const START_MUTATION = 'formula.mutation.set-formula-calculation-start'
const STOP_MUTATION = 'formula.mutation.set-formula-calculation-stop'
const NOTIFICATION_MUTATION = 'formula.mutation.set-formula-calculation-notification'

/** Chunks arrive tens of milliseconds apart; a scroll burst ends within this, so the merged cycle lands in the pause after it. */
export const FORMULA_STREAM_HOLD_MS = 1000

interface TriggerControllerInternals {
  _executingDirtyData?: Record<string, unknown>
  _executionInProgressParams?: unknown
  _restartCalculation?: boolean
}

interface HoldState {
  lastChunkAt: number
  vetoed: boolean
  inProgress: boolean
  timer: ReturnType<typeof setTimeout> | null
}

const states = new WeakMap<UniverRuntime, HoldState>()
let current: HoldState | null = null

export function hasDirtyData(dirty: Record<string, unknown> | undefined): boolean {
  if (!dirty) return false
  return Object.values(dirty).some((value) =>
    Array.isArray(value)
      ? value.length > 0
      : value && typeof value === 'object'
        ? Object.keys(value).length > 0
        : Boolean(value),
  )
}

/** Called by the importer after every file-data chunk lands in the grid. */
export function noteFormulaStreamChunk(): void {
  if (!current) return
  current.lastChunkAt = Date.now()
}

export function installFormulaStreamHold(runtime: UniverRuntime): void {
  if (states.has(runtime)) {
    current = states.get(runtime) ?? null
    return
  }
  const state: HoldState = { lastChunkAt: 0, vetoed: false, inProgress: false, timer: null }
  states.set(runtime, state)
  current = state
  const injector = runtime.univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const holding = () => Date.now() - state.lastChunkAt < FORMULA_STREAM_HOLD_MS

  const flush = () => {
    state.timer = null
    if (holding() || state.inProgress) {
      schedule()
      return
    }
    state.vetoed = false
    const controller = injector.get(TriggerCalculationController) as unknown as
      TriggerControllerInternals | undefined
    if (!controller) return
    // A restart the controller dispatched while held was vetoed, leaving its
    // "in progress" bookkeeping stale; clear it before the merged cycle.
    controller._executionInProgressParams = null
    controller._restartCalculation = false
    if (!hasDirtyData(controller._executingDirtyData)) return
    void commandService.executeCommand(
      START_MUTATION,
      { ...controller._executingDirtyData },
      {
        onlyLocal: true,
      },
    )
  }
  const schedule = () => {
    if (state.timer) clearTimeout(state.timer)
    state.timer = setTimeout(flush, FORMULA_STREAM_HOLD_MS + 20)
  }

  commandService.beforeCommandExecuted((command) => {
    if (command.id !== START_MUTATION || !holding()) return
    const params = command.params as { forceCalculation?: boolean } | undefined
    if (params?.forceCalculation) return
    // Same stack hygiene as the manual-calculation veto in calc-options.ts.
    const stack = (commandService as unknown as { _commandExecutionStack?: unknown[] })
      ._commandExecutionStack
    const index = stack?.indexOf(command) ?? -1
    if (index >= 0) stack?.splice(index, 1)
    state.vetoed = true
    schedule()
    throw new CustomCommandExecutionError('formula calculation held while file data streams in')
  })
  commandService.onCommandExecuted((command) => {
    if (command.id === STOP_MUTATION) state.inProgress = false
    if (command.id !== NOTIFICATION_MUTATION) return
    const params = command.params as { stageInfo?: unknown } | undefined
    state.inProgress = params?.stageInfo != null
  })
}
