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
