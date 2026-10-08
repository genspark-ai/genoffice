/**
 * A cell edit committed while the save IPC is in flight is not in the
 * request; the reopened session must carry it as a pending edit instead of
 * dropping it with the old session.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { handleSave, type SaveContext } from '../src/renderer/save-actions'
import {
  carryLateJournalCells,
  createEditJournal,
  recordSetRangeValues,
  recordStructuralOp,
  snapshotJournalCells,
  type EditJournal,
} from '../src/renderer/edit-journal'

const saveWorkbookEdits = vi.fn()
const writeWorkbookRecovery = vi.fn()

beforeEach(() => {
  saveWorkbookEdits.mockReset()
  writeWorkbookRecovery.mockReset().mockResolvedValue({ ok: true })
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: { saveWorkbookEdits, writeWorkbookRecovery },
  }
})

function stateWith(journal: EditJournal, sessionId: string) {
  return {
    editJournal: journal,
    recalc: {
      timer: null,
      generation: 0,
      failed: false,
      formulaCells: new Map(),
      overlay: new Map(),
    },
    flags: { preloadComplete: true },
    file: { sessionId, needsSaveAs: false, restoredFromRecovery: false },
  }
}

describe('carryLateJournalCells', () => {
  it('copies only entries recorded after the snapshot', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'saved' } } })
    const snapshot = snapshotJournalCells(journal)
    recordSetRangeValues(journal, 'sheet-1', { 1: { 0: { v: 'late' } } })
    recordSetRangeValues(journal, 'sheet-2', { 0: { 0: { v: 'late too' } } })
    const into = createEditJournal()
    expect(carryLateJournalCells(journal, snapshot, into)).toBe(2)
    expect(into.cells.get('sheet-1')?.get('0:0')).toBeUndefined()
    expect(into.cells.get('sheet-1')?.get('1:0')?.value).toBe('late')
    expect(into.cells.get('sheet-2')?.get('0:0')?.value).toBe('late too')
  })

  it('a re-edit of an already saved cell carries the merged entry', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'saved' } } })
    const snapshot = snapshotJournalCells(journal)
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'changed again' } } })
    const into = createEditJournal()
    expect(carryLateJournalCells(journal, snapshot, into)).toBe(1)
    expect(into.cells.get('sheet-1')?.get('0:0')?.value).toBe('changed again')
  })

  it('skips a sheet whose rows shifted after the snapshot', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 5: { 0: { v: 'saved' } } })
    recordSetRangeValues(journal, 'sheet-2', { 0: { 0: { v: 'saved' } } })
    const snapshot = snapshotJournalCells(journal)
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 0, count: 2 })
    recordSetRangeValues(journal, 'sheet-2', { 1: { 0: { v: 'late' } } })
    const into = createEditJournal()
    expect(carryLateJournalCells(journal, snapshot, into)).toBe(1)
    expect(into.cells.has('sheet-1')).toBe(false)
    expect(into.cells.get('sheet-2')?.get('1:0')?.value).toBe('late')
  })

  it('carries nothing when the journal did not change', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'saved' } } })
    const into = createEditJournal()
    expect(carryLateJournalCells(journal, snapshotJournalCells(journal), into)).toBe(0)
    expect(into.cells.size).toBe(0)
  })
})

describe('handleSave keeps edits made while the save IPC is pending', () => {
  function run(editDuringSave: boolean) {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'saved' } } })
    const lazyWorkbookRef = { current: stateWith(journal, '11111111-1111-4111-8111-111111111111') }
    const reopened = stateWith(createEditJournal(), '22222222-2222-4222-8222-222222222222')
    const file = { sessionId: reopened.file.sessionId, path: '/tmp/a.xlsx', sha256: 'b' }
    saveWorkbookEdits.mockImplementation(async () => {
      if (editDuringSave) {
        recordSetRangeValues(journal, 'sheet-1', { 1: { 0: { v: 'late' } } })
      }
      return { canceled: false, file }
    })
    const setPendingEdits = vi.fn()
    const ctx: SaveContext = {
      univerRef: { current: null },
      stashViewRestore: () => {},
      lazyWorkbookRef: lazyWorkbookRef as never,
      setMessage: () => {},
      setPendingEdits,
      openLazyWorkbook: vi.fn(async () => {
        lazyWorkbookRef.current = reopened
        setPendingEdits(0)
        return true
      }),
    }
    return { ctx, reopened, setPendingEdits }
  }

  it('carries the late cell into the reopened journal and marks it pending', async () => {
    const { ctx, reopened, setPendingEdits } = run(true)
    await expect(handleSave(ctx, 'save')).resolves.toMatchObject({ ok: true })
    expect(reopened.editJournal.cells.get('sheet-1')?.get('1:0')?.value).toBe('late')
    expect(reopened.editJournal.cells.get('sheet-1')?.has('0:0')).toBe(false)
    expect(setPendingEdits).toHaveBeenLastCalledWith(1)
  })

  it('leaves the reopened journal empty when nothing changed during the save', async () => {
    const { ctx, reopened, setPendingEdits } = run(false)
    await expect(handleSave(ctx, 'save')).resolves.toMatchObject({ ok: true })
    expect(reopened.editJournal.cells.size).toBe(0)
    expect(setPendingEdits).toHaveBeenLastCalledWith(0)
  })
})
