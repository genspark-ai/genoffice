/**
 * Shared gate for the two periodic workbook savers in App.tsx: the AutoSave
 * tick ('save', every 30s and on window blur) and the crash-recovery copy
 * tick ('recovery', every 30s). They used to carry independent in-flight
 * flags, so with AutoSave on, a dirty workbook could start BOTH saves
 * concurrently; in the main process the first finisher tears the workbook
 * session down while the second is still reading, and the survivor reports
 * "Unknown workbook session." (save failed) although the file was written.
 * One gate: at most one save at a time. With AutoSave on, its real save
 * flushes the journal so the recovery tick no-ops; with AutoSave off, the
 * 30s crash-recovery guarantee is unchanged. Pure so it can be unit tested.
 */
export interface SaveTickState {
  /** a save of either kind is currently running */
  saveInFlight: boolean
  /** no workbook open */
  hasWorkbook: boolean
  /** nothing unsaved since the last flush */
  journalEmpty: boolean
  /**
   * The reader's marks differ from what the file was opened with.
   *
   * A mark is not in the journal, so a workbook whose only edit is a
   * newly withheld cell would otherwise look clean to both ticks: no
   * AutoSave, no crash-recovery copy, and the mark lost when the tab
   * closes — with the reader having seen it drawn on the grid.
   */
  redactionsPendingSave: boolean
  /** the in-cell editor is open (its pending text is not in the journal yet; a save reloads the workbook and would wipe the edit) */
  editingCell: boolean
  /** converted .xls import whose first save must open a Save As dialog (a new unsaved workbook saves its backing file quietly instead) */
  needsSaveAsNotUnsavedNew: boolean
  /** CSV sessions: AutoSave would silently flatten the user's file */
  isCsv: boolean
  kind: 'save' | 'recovery'
  /** recovery-only: the session is backed by the recovery copy itself */
  restoredFromRecovery: boolean
  /** recovery-only: the user declined recovery for this workbook */
  automaticRecoveryDisabled: boolean
}

export function shouldRunSaveTick(s: SaveTickState): boolean {
  if (s.saveInFlight || !s.hasWorkbook) return false
  if (s.journalEmpty && !s.redactionsPendingSave) return false
  if (s.editingCell || s.needsSaveAsNotUnsavedNew || s.isCsv) return false
  if (s.kind === 'recovery' && (s.restoredFromRecovery || s.automaticRecoveryDisabled)) return false
  return true
}
