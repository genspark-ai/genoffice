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
import { shouldRunSaveTick } from '../src/renderer/save-scheduler'

function dirtyWorkbookState(overrides: Partial<Parameters<typeof shouldRunSaveTick>[0]> = {}) {
  return {
    saveInFlight: false,
    hasWorkbook: true,
    journalEmpty: false,
    redactionsPendingSave: false,
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

describe('a workbook whose only edit is a withheld cell still saves', () => {
  /**
   * A mark lives in a package part, not in the edit journal, so the journal
   * looks empty. Before this was checked, such a workbook was skipped by both
   * ticks: no AutoSave, no crash-recovery copy, and the mark lost on close —
   * with the reader having watched it get drawn.
   */
  const onlyRedactions = { journalEmpty: true, redactionsPendingSave: true }

  it('runs the AutoSave tick', () => {
    expect(shouldRunSaveTick(dirtyWorkbookState({ ...onlyRedactions, kind: 'save' }))).toBe(true)
  })

  it('runs the crash-recovery tick', () => {
    expect(shouldRunSaveTick(dirtyWorkbookState({ ...onlyRedactions, kind: 'recovery' }))).toBe(
      true,
    )
  })

  it('still skips when nothing at all is pending', () => {
    // The flag is what makes the difference, so a workbook that is genuinely
    // clean must keep skipping — otherwise every idle workbook would save
    // every thirty seconds.
    expect(
      shouldRunSaveTick(
        dirtyWorkbookState({ journalEmpty: true, redactionsPendingSave: false, kind: 'save' }),
      ),
    ).toBe(false)
  })

  it('still honours the in-cell editor guard', () => {
    expect(
      shouldRunSaveTick(dirtyWorkbookState({ ...onlyRedactions, kind: 'save', editingCell: true })),
    ).toBe(false)
  })
})
