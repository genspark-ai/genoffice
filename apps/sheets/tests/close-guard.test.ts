/**
 * On SIGTERM the close prompt resolved to its default button ("Save")
 * with nobody answering, so a restart / installer / killall silently overwrote the
 * user's original workbook. The guard must proceed without writing while the app is
 * shutting down — unsaved work is covered by the periodic recovery copy.
 */
import { describe, expect, it } from 'vitest'
import { closeGuardDecision } from '../src/main/close-guard'

describe('closeGuardDecision', () => {
  it('prompts only when there are pending edits and the app is staying up', () => {
    expect(closeGuardDecision({ pendingEdits: 3, destroyed: false, shuttingDown: false })).toBe(
      'prompt',
    )
  })

  it('never prompts (and so never saves) while shutting down', () => {
    expect(closeGuardDecision({ pendingEdits: 3, destroyed: false, shuttingDown: true })).toBe(
      'proceed',
    )
  })

  it('a clean workbook or a dead renderer just proceeds', () => {
    expect(closeGuardDecision({ pendingEdits: 0, destroyed: false, shuttingDown: false })).toBe(
      'proceed',
    )
    expect(closeGuardDecision({ pendingEdits: 5, destroyed: true, shuttingDown: false })).toBe(
      'proceed',
    )
  })

  it('shutdown wins over every other input', () => {
    for (const pendingEdits of [0, 1, 9999]) {
      for (const destroyed of [false, true]) {
        expect(closeGuardDecision({ pendingEdits, destroyed, shuttingDown: true })).toBe('proceed')
      }
    }
  })
})
