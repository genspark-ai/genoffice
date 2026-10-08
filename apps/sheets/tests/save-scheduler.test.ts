/**
 * Regression: the AutoSave tick ('save') and the crash-recovery copy tick
 * ('recovery') each carried their own in-flight flag, so with AutoSave on a
 * dirty workbook started both saves concurrently every 30 s. In the main
 * process the first finisher tears the workbook session down while the second
 * is still reading; the survivor reported "Unknown workbook session."
 * (save failed) although the file had been written. Both ticks now share one
 * gate (save-scheduler.ts / saveInFlightRef), so a second tick skips instead
 * of racing, and retries on its next 30 s wake-up.
 */
import { describe, expect, it } from 'vitest'
import { createSaveGate, shouldRunSaveTick } from '../src/renderer/save-scheduler'

function dirtyWorkbookState(overrides: Partial<Parameters<typeof shouldRunSaveTick>[0]> = {}) {
  return {
    saveInFlight: false,
    hasWorkbook: true,
    journalEmpty: false,
    editingCell: false,
    needsSaveAsNotUnsavedNew: false,
    isCsv: false,
    kind: 'recovery' as const,
    restoredFromRecovery: false,
    automaticRecoveryDisabled: false,
    ...overrides,
  }
}

describe('shouldRunSaveTick', () => {
  it('lets a clean dirty-workbook tick of either kind run', () => {
    expect(shouldRunSaveTick(dirtyWorkbookState({ kind: 'save' }))).toBe(true)
    expect(shouldRunSaveTick(dirtyWorkbookState({ kind: 'recovery' }))).toBe(true)
  })

  it('the regression: while either save is in flight, the OTHER tick skips instead of racing', () => {
    // AutoSave started first → the recovery tick 30 s later must skip
    expect(shouldRunSaveTick(dirtyWorkbookState({ kind: 'recovery', saveInFlight: true }))).toBe(
      false,
    )
    // recovery copy started first (AutoSave off → on mid-session, or blur
    // firing the AutoSave tick early) → the AutoSave tick must skip
    expect(shouldRunSaveTick(dirtyWorkbookState({ kind: 'save', saveInFlight: true }))).toBe(false)
  })

  it('a skipped tick retries after the in-flight save finishes', () => {
    const inFlight = dirtyWorkbookState({ kind: 'save', saveInFlight: true })
    expect(shouldRunSaveTick(inFlight)).toBe(false)
    // the AutoSave .finally() clears the shared gate
    expect(shouldRunSaveTick({ ...inFlight, saveInFlight: false })).toBe(true)
  })

  it('an AutoSave real save makes the recovery tick no-op via the empty journal', () => {
    // after AutoSave flushed, the journal is clean: recovery has nothing to copy
    expect(shouldRunSaveTick(dirtyWorkbookState({ kind: 'recovery', journalEmpty: true }))).toBe(
      false,
    )
  })

  it('the 30 s crash-recovery guarantee holds with AutoSave off', () => {
    // AutoSave off: nothing else ever sets saveInFlight, so every dirty tick runs
    for (let i = 0; i < 3; i++) {
      expect(shouldRunSaveTick(dirtyWorkbookState({ kind: 'recovery' }))).toBe(true)
    }
  })

  it('keeps the pre-existing per-kind guards', () => {
    expect(shouldRunSaveTick(dirtyWorkbookState({ hasWorkbook: false }))).toBe(false)
    expect(shouldRunSaveTick(dirtyWorkbookState({ journalEmpty: true }))).toBe(false)
    expect(shouldRunSaveTick(dirtyWorkbookState({ editingCell: true }))).toBe(false)
    expect(shouldRunSaveTick(dirtyWorkbookState({ needsSaveAsNotUnsavedNew: true }))).toBe(false)
    expect(shouldRunSaveTick(dirtyWorkbookState({ isCsv: true }))).toBe(false)
    // recovery-only guards do not block the real save kind
    expect(shouldRunSaveTick(dirtyWorkbookState({ restoredFromRecovery: true }))).toBe(false)
    expect(shouldRunSaveTick(dirtyWorkbookState({ automaticRecoveryDisabled: true }))).toBe(false)
    expect(
      shouldRunSaveTick(dirtyWorkbookState({ kind: 'save', restoredFromRecovery: true })),
    ).toBe(true)
  })
})

const settle = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

describe('createSaveGate', () => {
  it('runs saves one after another and reports busy until the last one settles', async () => {
    const gate = createSaveGate()
    const order: string[] = []
    let releaseFirst!: () => void
    const first = gate.run(async () => {
      order.push('first:start')
      await new Promise<void>((resolve) => (releaseFirst = resolve))
      order.push('first:end')
      return 1
    })
    const second = gate.run(async () => {
      order.push('second:start')
      return 2
    })
    expect(gate.busy).toBe(true)
    await settle()
    expect(order).toEqual(['first:start'])
    releaseFirst()
    expect(await first).toBe(1)
    expect(await second).toBe(2)
    expect(order).toEqual(['first:start', 'first:end', 'second:start'])
    await settle()
    expect(gate.busy).toBe(false)
  })

  it('a failed save does not block the next one and the gate reopens', async () => {
    const gate = createSaveGate()
    const failed = gate.run(() => Promise.reject(new Error('boom')))
    await expect(failed).rejects.toThrow('boom')
    expect(await gate.run(async () => 'ok')).toBe('ok')
    await settle()
    expect(gate.busy).toBe(false)
  })
})
