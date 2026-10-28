/**
 * In-renderer notice that files on disk changed through the chrome (a tab
 * rename, for instance). Home and the tab strip live in the same renderer but
 * do not share state, and Home otherwise re-pulls only on window focus — which
 * does not fire when the shell webContents already held focus (the rename
 * input lives in it), so a rename from the tab strip would leave the Home list
 * showing the old name until the next focus change.
 */
const target = new EventTarget()
const FILES_CHANGED = 'files-changed'

export function notifyFilesChanged(): void {
  target.dispatchEvent(new Event(FILES_CHANGED))
}

/** Subscribe; returns the unsubscribe function (effect-cleanup friendly). */
export function onFilesChanged(handler: () => void): () => void {
  target.addEventListener(FILES_CHANGED, handler)
  return () => target.removeEventListener(FILES_CHANGED, handler)
}

/**
 * Main → renderer half: the main process pushes HOME_CHANNELS.recentsChanged
 * whenever a save or open lands a new head entry in the recents list (the docs
 * pushRecent funnel broadcasts it). Bridged onto the event above so Home
 * re-pulls its list the moment a save finishes — a document saved in an editor
 * tab never moves window focus, so the focus listener alone would leave the new
 * file missing from Home until the next focus change.
 */
export function bindRecentsChanged(): () => void {
  // optional-chained: component tests stub a partial aiOffice, and a preload
  // without the method degrades to the old focus-driven refresh
  const off = window.aiOffice.onRecentsChanged?.(() => notifyFilesChanged())
  return () => off?.()
}
