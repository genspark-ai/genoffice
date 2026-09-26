/**
 * On SIGTERM the close prompt resolved to its default button ("Save")
 * with nobody answering, so a restart / installer / killall silently overwrote the
 * user's original workbook. The guard must proceed without writing while the app is
 * shutting down — unsaved work is covered by the periodic recovery copy.
 */
import { describe, expect, it, vi } from 'vitest'
import { closeGuardDecision } from '../src/main/close-guard'
import {
  commitActiveCellEditor,
  pendingEditsForClose,
  runAfterCellEditorCommit,
  type ActiveCellEditor,
} from '../src/renderer/univer-state'

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

describe('pendingEditsForClose', () => {
  it('counts a dirty in-cell editor even when the journal is empty', () => {
    expect(pendingEditsForClose(0, true)).toBe(1)
    expect(pendingEditsForClose(3, true)).toBe(4)
    expect(pendingEditsForClose(3, false)).toBe(3)
  })
})

describe('runAfterCellEditorCommit', () => {
  it('commits before a save, reopen, or replacement action', async () => {
    const calls: string[] = []
    const editor: ActiveCellEditor = {
      isCellEditing: () => true,
      endEditingAsync: async () => {
        calls.push('commit')
        return true
      },
    }

    await expect(
      runAfterCellEditorCommit(editor, () => {
        calls.push('transition')
      }),
    ).resolves.toBe(true)
    expect(calls).toEqual(['commit', 'transition'])
  })

  it('does not run a transition after a failed commit', async () => {
    const transition = vi.fn()
    const editor: ActiveCellEditor = {
      isCellEditing: () => true,
      endEditingAsync: async () => false,
    }

    await expect(runAfterCellEditorCommit(editor, transition)).resolves.toBe(false)
    expect(transition).not.toHaveBeenCalled()
  })
})

describe('commitActiveCellEditor', () => {
  it('commits the active editor before returning', async () => {
    const endEditingAsync = vi.fn(async () => true)
    const editor: ActiveCellEditor = {
      isCellEditing: () => true,
      endEditingAsync,
    }

    await expect(commitActiveCellEditor(editor)).resolves.toBe(true)
    expect(endEditingAsync).toHaveBeenCalledWith(true)
  })

  it('does nothing when no cell editor is active', async () => {
    const endEditingAsync = vi.fn(async () => true)
    const editor: ActiveCellEditor = {
      isCellEditing: () => false,
      endEditingAsync,
    }

    await expect(commitActiveCellEditor(editor)).resolves.toBe(true)
    await expect(commitActiveCellEditor(null)).resolves.toBe(true)
    expect(endEditingAsync).not.toHaveBeenCalled()
  })

  it('reports a failed commit so close cannot continue', async () => {
    const editor: ActiveCellEditor = {
      isCellEditing: () => true,
      endEditingAsync: vi.fn(async () => false),
    }

    await expect(commitActiveCellEditor(editor)).resolves.toBe(false)
  })
})
