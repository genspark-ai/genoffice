import { describe, expect, it } from 'vitest'

import {
  calculateSheet,
  isManualCalculation,
  resetCalculationMode,
  setManualCalculation,
} from '../src/renderer/calc-options'
import type { UniverRuntime } from '../src/renderer/univer-state'

const START = 'formula.mutation.set-formula-calculation-start'

interface CommandInfo {
  id: string
  params?: Record<string, unknown> | undefined
}

function makeRuntime() {
  const listeners: ((command: CommandInfo) => void)[] = []
  const executed: CommandInfo[] = []
  // Mirrors Univer's CommandService: the entry is pushed before the
  // before-hooks run and removed by identity only on the success path — a
  // hook that throws leaves it stranded.
  const stack: CommandInfo[] = []
  const commandService = {
    _commandExecutionStack: stack,
    beforeCommandExecuted: (listener: (command: CommandInfo) => void) => {
      listeners.push(listener)
    },
    syncExecuteCommand: (id: string, params?: Record<string, unknown>) => {
      const info = { id, params }
      stack.push(info)
      try {
        for (const listener of listeners) listener(info)
      } catch {
        return false
      }
      executed.push(info)
      stack.splice(stack.indexOf(info), 1)
      return true
    },
  }
  const sheet = { getRowCount: () => 100, getColumnCount: () => 10 }
  const worksheet = {
    getSheet: () => sheet,
    getSheetId: () => 's1',
    getSheetName: () => 'Sheet1',
  }
  const workbook = { getActiveSheet: () => worksheet, getId: () => 'wb' }
  const runtime = {
    univer: { __getInjector: () => ({ get: () => commandService }) },
    univerAPI: { getActiveWorkbook: () => workbook },
  } as unknown as UniverRuntime
  const engineRecalc = (params: Record<string, unknown>) =>
    commandService.syncExecuteCommand(START, params)
  return { runtime, executed, stack, engineRecalc, listenerCount: () => listeners.length }
}

describe('calculation options', () => {
  it('manual mode cancels the recalc without stranding a stack entry', () => {
    const { runtime, engineRecalc, executed, stack } = makeRuntime()
    setManualCalculation(runtime, true)
    const result = engineRecalc({ commands: [], dirtyRanges: [{ unitId: 'wb' }] })
    // No calculation cycle at all — an emptied-but-executed mutation would
    // still refresh volatile cells and fire calculationEnd.
    expect(result).toBe(false)
    expect(executed).toHaveLength(0)
    // Univer only cleans the execution stack on the success path; the hook
    // removes its own entry before throwing, so the stack is balanced the
    // moment the veto lands — no deferred-cleanup window.
    expect(stack).toHaveLength(0)
  })

  it('scrubs by identity, leaving a parent command in flight untouched', () => {
    const { runtime, engineRecalc, stack } = makeRuntime()
    setManualCalculation(runtime, true)
    // A cell edit is mid-dispatch when its recalc kick-off gets vetoed.
    const parent: CommandInfo = { id: 'sheet.command.set-range-values' }
    stack.push(parent)
    engineRecalc({ dirtyRanges: [{ unitId: 'wb' }] })
    expect(stack).toEqual([parent])
  })

  it('lets forced recalculation through in manual mode', () => {
    const { runtime, engineRecalc, executed } = makeRuntime()
    setManualCalculation(runtime, true)
    engineRecalc({ forceCalculation: true, dirtyRanges: [{ unitId: 'wb' }] })
    expect(executed).toHaveLength(1)
  })

  it('does not intervene in automatic mode', () => {
    const { runtime, engineRecalc, executed, stack } = makeRuntime()
    setManualCalculation(runtime, false)
    engineRecalc({ dirtyRanges: [{ unitId: 'wb' }] })
    expect(executed).toHaveLength(1)
    expect(stack).toHaveLength(0)
  })

  it('calculate-sheet marks the whole sheet dirty and bypasses the veto', () => {
    const { runtime, executed } = makeRuntime()
    setManualCalculation(runtime, true)
    calculateSheet(runtime)
    expect(executed).toHaveLength(1)
    const params = executed[0]?.params as {
      dirtyRanges: { range: { endRow: number; endColumn: number } }[]
      dirtyNameMap: Record<string, Record<string, string>>
    }
    // dirtyNameMap dirties every formula on the sheet (Shift+F9 semantics,
    // cross-sheet precedents included); dirtyRanges covers dependents.
    expect(params.dirtyNameMap).toEqual({ wb: { s1: 'Sheet1' } })
    expect(params.dirtyRanges[0]?.range).toEqual({
      startRow: 0,
      endRow: 99,
      startColumn: 0,
      endColumn: 9,
    })
  })

  it('resets to automatic for the next workbook', () => {
    const { runtime, engineRecalc } = makeRuntime()
    setManualCalculation(runtime, true)
    expect(isManualCalculation(runtime)).toBe(true)
    resetCalculationMode(runtime)
    expect(isManualCalculation(runtime)).toBe(false)
    const params: Record<string, unknown> = { dirtyRanges: [{ unitId: 'wb' }] }
    engineRecalc(params)
    expect(params.dirtyRanges).toEqual([{ unitId: 'wb' }])
  })

  it('keeps state per runtime and installs one hook each', () => {
    const first = makeRuntime()
    const second = makeRuntime()
    setManualCalculation(first.runtime, true)
    expect(isManualCalculation(first.runtime)).toBe(true)
    expect(isManualCalculation(second.runtime)).toBe(false)
    setManualCalculation(first.runtime, false)
    setManualCalculation(first.runtime, true)
    expect(first.listenerCount()).toBe(1)
  })
})
