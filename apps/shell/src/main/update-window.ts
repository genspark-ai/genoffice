import { join } from 'node:path'
import { BrowserWindow, ipcMain } from 'electron'
import type { UpdateUiState } from '../shared/update-api'
import { UPDATE_CHANNELS } from '../shared/update-api'

/**
 * The strong-guidance update window: a frameless modal
 * card centered over the shell window. Content lives in renderer/update.html;
 * this module owns the window lifecycle, the pushed UI state, and the
 * download / install / later IPC surface.
 */

interface UpdateActions {
  onDownload: () => void
  onInstall: () => void
  onLater: () => void
  onOpenDownload: () => void
}

let updateWin: BrowserWindow | null = null
let currentState: UpdateUiState | null = null
let actions: UpdateActions | null = null
let ipcRegistered = false
// distinguishes programmatic close (install/quit) from the user closing the
// window some other way (Alt+F4…), which counts as "later"
let closingProgrammatically = false

function registerIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true
  ipcMain.handle(UPDATE_CHANNELS.getState, () => currentState)
  ipcMain.handle(UPDATE_CHANNELS.download, () => actions?.onDownload())
  ipcMain.handle(UPDATE_CHANNELS.install, () => actions?.onInstall())
  ipcMain.handle(UPDATE_CHANNELS.later, () => actions?.onLater())
  ipcMain.handle(UPDATE_CHANNELS.openDownload, () => actions?.onOpenDownload())
}

export function showUpdateWindow(
  parent: BrowserWindow | null,
  state: UpdateUiState,
  windowActions: UpdateActions,
): void {
  currentState = state
  actions = windowActions
  registerIpc()

  if (updateWin && !updateWin.isDestroyed()) {
    updateWin.webContents.send(UPDATE_CHANNELS.changed, currentState)
    updateWin.focus()
    return
  }

  const win = new BrowserWindow({
    width: 400,
    height: 430,
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
      preload: join(__dirname, '../preload/update.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  updateWin = win

  win.once('ready-to-show', () => win.show())
  win.on('closed', () => {
    updateWin = null
    if (!closingProgrammatically) actions?.onLater()
    closingProgrammatically = false
  })

  if (process.env.ELECTRON_RENDERER_URL) {
    void win.loadURL(`${process.env.ELECTRON_RENDERER_URL}/update.html`)
  } else {
    void win.loadFile(join(__dirname, '../renderer/update.html'))
  }
}

export function isUpdateWindowOpen(): boolean {
  return updateWin !== null && !updateWin.isDestroyed()
}

export function pushUpdateState(patch: Partial<UpdateUiState>): void {
  if (!currentState) return
  currentState = { ...currentState, ...patch }
  if (updateWin && !updateWin.isDestroyed()) {
    updateWin.webContents.send(UPDATE_CHANNELS.changed, currentState)
  }
}

export function closeUpdateWindow(): void {
  if (updateWin && !updateWin.isDestroyed()) {
    closingProgrammatically = true
    updateWin.close()
  }
  updateWin = null
}
