/**
 * Close-guard decision, kept free of Electron so it can be reasoned about (and
 * tested) on its own.
 *
 * 'proceed' lets the window go without writing anything; 'prompt' asks the user
 * Save / Don't Save / Cancel. Shutdown must always be 'proceed': a dialog raised
 * while the app is quitting resolves to its default button, and with Save as the
 * default that silently overwrote the user's original file on SIGTERM — a restart,
 * an installer, or killall could rewrite a model nobody chose to save.
 * Unsaved work is covered by the periodic recovery copy, which the next launch
 * offers to restore.
 */
export function closeGuardDecision(state: {
  /** Pending edits the renderer last reported */
  pendingEdits: number
  /** The renderer is already gone */
  destroyed: boolean
  /** before-quit / SIGTERM / SIGINT seen */
  shuttingDown: boolean
}): 'proceed' | 'prompt' {
  if (state.pendingEdits <= 0) return 'proceed'
  if (state.destroyed) return 'proceed'
  if (state.shuttingDown) return 'proceed'
  return 'prompt'
}
