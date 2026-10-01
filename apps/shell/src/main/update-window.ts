import { join } from 'node:path'
import { BrowserWindow, ipcMain } from 'electron'
import type { UpdateUiState } from '../shared/update-api'
import { UPDATE_CHANNELS } from '../shared/update-api'

/**
 * The update window: a frameless card centered over the shell window —
 * non-modal since the settings entry landed, so it never blocks work; its
 * "minimize" action folds it away and Settings → About re-opens it on demand.
 * Content lives in renderer/update.html; this module owns the window
 * lifecycle, the pushed UI state, and the download / install / later IPC
 * surface, plus the settings-facing open-for-update entry.
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
/** the shell window the card belongs to; also the state-changed listener */
let lastParent: BrowserWindow | null = null
// distinguishes programmatic close (install/quit) from the user closing the
// window some other way (Alt+F4…), which counts as "later"
let closingProgrammatically = false

function broadcastState(): void {
  if (lastParent && !lastParent.isDestroyed())
    lastParent.webContents.send(UPDATE_CHANNELS.stateChanged, currentState)
}

function registerIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true
  ipcMain.handle(UPDATE_CHANNELS.getState, () => currentState)
  ipcMain.handle(UPDATE_CHANNELS.download, () => actions?.onDownload())
  ipcMain.handle(UPDATE_CHANNELS.install, () => actions?.onInstall())
  ipcMain.handle(UPDATE_CHANNELS.later, () => actions?.onLater())
  ipcMain.handle(UPDATE_CHANNELS.openDownload, () => actions?.onOpenDownload())
  // Settings → About "update to vX": surface the dialog again (it may have
  // been minimized) and, when nothing has started yet, start the download.
  // Returns false while no update is known — the settings button stays
  // hidden in that state, so callers treat it as a no-op.
  ipcMain.handle(UPDATE_CHANNELS.openForUpdate, (): boolean => {
    if (!currentState) return false
    if (updateWin && !updateWin.isDestroyed()) {
      updateWin.show()
      updateWin.focus()
    } else if (lastParent && !lastParent.isDestroyed()) {
      showUpdateWindow(lastParent, currentState, actions ?? dummyActions)
    }
    if (currentState.phase === 'available' || currentState.phase === 'error') actions?.onDownload()
    return true
  })
}

/** keeps the handler total when actions are missing (cannot happen today) */
const dummyActions: UpdateActions = {
  onDownload: () => {},
  onInstall: () => {},
  onLater: () => {},
  onOpenDownload: () => {},
}

/** the freshest pushed state (null before the first update was ever seen) */
export function currentUpdateUiState(): UpdateUiState | null {
  return currentState
}

export function showUpdateWindow(
  parent: BrowserWindow | null,
  state: UpdateUiState,
  windowActions: UpdateActions,
): void {
  currentState = state
  actions = windowActions
  if (parent && !parent.isDestroyed()) lastParent = parent
  registerIpc()
  broadcastState()

  if (updateWin && !updateWin.isDestroyed()) {
    updateWin.webContents.send(UPDATE_CHANNELS.changed, currentState)
    updateWin.show()
    updateWin.focus()
    return
  }

  const win = new BrowserWindow({
    width: 400,
    height: 430,
    ...(parent && !parent.isDestroyed() ? { parent } : {}),
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
  broadcastState()
}

export function closeUpdateWindow(): void {
  if (updateWin && !updateWin.isDestroyed()) {
    closingProgrammatically = true
    updateWin.close()
  }
  updateWin = null
}
