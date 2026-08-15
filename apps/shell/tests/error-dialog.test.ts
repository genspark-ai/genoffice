import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { BrowserWindow } from 'electron'

/**
 * showErrorDialog (src/main/error-dialog.ts): surfacing a create-tab failure
 * must never block the main process. dialog.showErrorBox is synchronous on
 * Windows (parentless TaskDialog): with the box unnoticed behind the shell
 * window, main-process JS stays parked in its nested pump, so every click,
 * native menu item and the window close button die with no "(Not Responding)"
 * — the alpha row-51 full-UI wedge. These tests simulate that scenario: an
 * error dialog nobody ever dismisses while failures keep coming.
 */

const showMessageBox = vi.fn<(...args: unknown[]) => Promise<unknown>>()
const showErrorBox = vi.fn()

vi.mock('electron', () => ({
  dialog: {
    showMessageBox: (...args: unknown[]) => showMessageBox(...args),
    showErrorBox: (...args: unknown[]) => showErrorBox(...args),
  },
}))

function fakeWindow(destroyed = false): BrowserWindow {
  return { isDestroyed: () => destroyed } as unknown as BrowserWindow
}

let showErrorDialog: typeof import('../src/main/error-dialog').showErrorDialog

beforeEach(async () => {
  vi.resetModules()
  showMessageBox.mockReset()
  showErrorBox.mockReset()
  showMessageBox.mockImplementation(() => new Promise(() => {}))
  showErrorDialog = (await import('../src/main/error-dialog')).showErrorDialog
})

describe('showErrorDialog', () => {
  it('uses the async parented message box, never the blocking showErrorBox', () => {
    const win = fakeWindow()
    showErrorDialog(win, 'Could not create the new document', new Error('spawn EACCES'))
    expect(showErrorBox).not.toHaveBeenCalled()
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    expect(showMessageBox).toHaveBeenCalledWith(win, {
      type: 'error',
      message: 'Could not create the new document',
      detail: 'spawn EACCES',
    })
  })

  it('returns synchronously even though the dialog is never dismissed', () => {
    // A hung dialog promise must not propagate: the caller (ipc handler /
    // menu click) has to finish so later input keeps being processed.
    const before = Date.now()
    showErrorDialog(fakeWindow(), 'boom', new Error('x'))
    expect(Date.now() - before).toBeLessThan(50)
  })

  it('does not stack dialogs while one is still open (repeated create failures)', async () => {
    let dismiss!: (v: unknown) => void
    showMessageBox.mockImplementation(() => new Promise((resolve) => (dismiss = resolve)))
    const win = fakeWindow()
    // row-51 sequence: AI Sheets, File > New > AI Sheets, AI Slides, AI Docs
    showErrorDialog(win, 'fail', new Error('1'))
    showErrorDialog(win, 'fail', new Error('2'))
    showErrorDialog(win, 'fail', new Error('3'))
    showErrorDialog(win, 'fail', new Error('4'))
    expect(showMessageBox).toHaveBeenCalledTimes(1)

    dismiss({ response: 0 })
    await Promise.resolve()
    showErrorDialog(win, 'fail', new Error('5'))
    expect(showMessageBox).toHaveBeenCalledTimes(2)
  })

  it('falls back to an unparented async dialog when the shell window is gone', async () => {
    showMessageBox.mockImplementation(() => Promise.resolve({ response: 0 }))
    showErrorDialog(null, 'fail', new Error('no window'))
    await Promise.resolve()
    showErrorDialog(fakeWindow(true), 'fail', 'plain string')
    expect(showMessageBox).toHaveBeenNthCalledWith(1, {
      type: 'error',
      message: 'fail',
      detail: 'no window',
    })
    expect(showMessageBox).toHaveBeenNthCalledWith(2, {
      type: 'error',
      message: 'fail',
      detail: 'plain string',
    })
    expect(showErrorBox).not.toHaveBeenCalled()
  })
})
