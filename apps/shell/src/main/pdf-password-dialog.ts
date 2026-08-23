import { join } from 'node:path'
import { BrowserWindow, ipcMain } from 'electron'
import type { PdfPasswordUiState } from '../shared/pdf-password-api'
import { PDF_PASSWORD_CHANNELS } from '../shared/pdf-password-api'

/**
 * Password prompt for encrypted PDFs (P23): a frameless modal card centered
 * over the shell window, same lifecycle pattern as update-window.ts. One
 * promptPdfPassword() call resolves with one attempt (the entered password,
 * or null when the user cancels / closes the window). The window survives
 * across attempts — a rejected password re-prompts in place via pushed state
 * — and the caller closes it with closePdfPasswordDialog() when the retry
 * loop ends (success, cancel or unrelated failure).
 */

let pwWin: BrowserWindow | null = null
let currentState: PdfPasswordUiState | null = null
let pendingResolve: ((password: string | null) => void) | null = null
let ipcRegistered = false

function settle(password: string | null): void {
  const resolve = pendingResolve
  pendingResolve = null
  resolve?.(password)
}

function registerIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true
  ipcMain.handle(PDF_PASSWORD_CHANNELS.getState, () => currentState)
  ipcMain.handle(PDF_PASSWORD_CHANNELS.submit, (_event, password: unknown) => {
    if (typeof password !== 'string') return
    pushState({ busy: true })
    settle(password)
  })
  ipcMain.handle(PDF_PASSWORD_CHANNELS.cancel, () => {
    settle(null)
    closePdfPasswordDialog()
  })
}

function pushState(patch: Partial<PdfPasswordUiState>): void {
  if (!currentState) return
  currentState = { ...currentState, ...patch }
  if (pwWin && !pwWin.isDestroyed()) {
    pwWin.webContents.send(PDF_PASSWORD_CHANNELS.changed, currentState)
  }
}

/** Ask the user for the PDF password; one call = one attempt. */
export function promptPdfPassword(
  parent: BrowserWindow | null,
  state: PdfPasswordUiState,
): Promise<string | null> {
  // a second prompt while one is pending would orphan the first caller —
  // resolve it as cancelled first (cannot happen in the single-flight export
  // path, but keeps the module safe on its own)
  settle(null)
  currentState = { ...state, busy: false }
  registerIpc()

  const promise = new Promise<string | null>((resolve) => {
    pendingResolve = resolve
  })

  if (pwWin && !pwWin.isDestroyed()) {
    pwWin.webContents.send(PDF_PASSWORD_CHANNELS.changed, currentState)
    pwWin.focus()
    return promise
  }

  const win = new BrowserWindow({
    width: 400,
    height: 260,
    ...(parent && !parent.isDestroyed() ? { parent, modal: true } : {}),
    frame: false,
    transparent: true,
    backgroundColor: '#00000000',
    resizable: false,
    minimizable: false,
    maximizable: false,
    fullscreenable: false,
    show: false,
    title: state.strings.title,
    webPreferences: {
      preload: join(__dirname, '../preload/pdf-password.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  pwWin = win

  win.once('ready-to-show', () => win.show())
  // closing the window any other way (Alt+F4…) counts as cancel
  win.on('closed', () => {
    pwWin = null
    settle(null)
  })

  if (process.env.ELECTRON_RENDERER_URL) {
    void win.loadURL(`${process.env.ELECTRON_RENDERER_URL}/pdf-password.html`)
  } else {
    void win.loadFile(join(__dirname, '../renderer/pdf-password.html'))
  }
  return promise
}

export function closePdfPasswordDialog(): void {
  if (pwWin && !pwWin.isDestroyed()) pwWin.close()
  pwWin = null
  currentState = null
}
