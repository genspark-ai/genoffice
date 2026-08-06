import { dialog } from 'electron'
import type { BrowserWindow } from 'electron'

let showing = false

/**
 * Never dialog.showErrorBox here: on Windows it blocks main-process JS in a
 * nested native pump and, parentless, can hide behind the window — a wedged
 * UI with no "(Not Responding)". Async + parented avoids both.
 */
export function showErrorDialog(win: BrowserWindow | null, message: string, err: unknown): void {
  if (showing) return
  showing = true
  const options = {
    type: 'error' as const,
    message,
    detail: err instanceof Error ? err.message : String(err),
  }
  const shown =
    win && !win.isDestroyed() ? dialog.showMessageBox(win, options) : dialog.showMessageBox(options)
  void shown.finally(() => {
    showing = false
  })
}
