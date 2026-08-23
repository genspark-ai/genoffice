/**
 * Presenter-view multi-screen show, extracted from slides-main.ts: audience
 * window on the external display + presenter->audience state forwarding.
 * The audience window shares the same Session as the presenter (sessions holds
 * one reference per wc.id), so read-only IPC (get-render-slides/get-animations/
 * get-transition…) works naturally, and it sees the in-memory document
 * (including unsaved changes) without re-reading from disk.
 */
import { BrowserWindow, ipcMain, screen } from 'electron'
import type { WebContents } from 'electron'
import type { AudienceNavAction, ShowInkEvent, ShowSyncState } from '../shared/ipc'
import { runtime, sessions, viewerWcIds, windowRefs } from './session-state'

interface PresenterShow {
  presenterWc: WebContents
  audienceWin: BrowserWindow | null
  /** Most recent sync state: audience-ready re-sends it when the audience window mounts after the broadcast */
  lastSync: ShowSyncState | null
}
const presenterShows = new Map<number, PresenterShow>()
/** Audience wc.id -> presenter wc.id (routing for navigation actions sent back) */
const audiencePresenter = new Map<number, number>()

/** Host window of the presenter renderer (in tab mode, the shell's single window) */
function presenterHostWindow(wc: WebContents): BrowserWindow | null {
  return BrowserWindow.fromWebContents(wc) ?? windowRefs.shellWindow
}

/** Fullscreen a window on a given display. macOS uses simpleFullScreen to avoid Spaces animations and cross-screen move restrictions */
function fullScreenOnDisplay(win: BrowserWindow, bounds: Electron.Rectangle): void {
  if (process.platform === 'darwin') {
    if (win.isSimpleFullScreen()) win.setSimpleFullScreen(false)
    win.setBounds(bounds)
    win.setSimpleFullScreen(true)
  } else {
    if (win.isFullScreen()) win.setFullScreen(false)
    win.setBounds(bounds)
    win.setFullScreen(true)
  }
}

function closePresenterShow(presenterId: number): void {
  const show = presenterShows.get(presenterId)
  if (!show) return
  presenterShows.delete(presenterId)
  const win = show.audienceWin
  show.audienceWin = null
  if (win && !win.isDestroyed()) {
    if (process.platform === 'darwin' && win.isSimpleFullScreen()) win.setSimpleFullScreen(false)
    win.close()
  }
}

/** Register the slides:presenter-* / slides:audience-* channels (called from registerSlidesIpc). */
export function registerPresenterIpc(): void {
  ipcMain.handle('slides:presenter-start', (e) => {
    const existing = presenterShows.get(e.sender.id)
    if (existing) return { audience: existing.audienceWin != null }
    const show: PresenterShow = { presenterWc: e.sender, audienceWin: null, lastSync: null }
    presenterShows.set(e.sender.id, show)
    e.sender.once('destroyed', () => closePresenterShow(e.sender.id))

    const session = sessions.get(e.sender.id)
    const host = presenterHostWindow(e.sender)
    const hostDisplay = host
      ? screen.getDisplayMatching(host.getBounds())
      : screen.getPrimaryDisplay()
    const external = screen.getAllDisplays().find((d) => d.id !== hostDisplay.id)
    if (!external || !session) return { audience: false }

    const win = new BrowserWindow({
      ...external.bounds,
      frame: false,
      show: false,
      backgroundColor: '#000000',
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
      },
    })
    show.audienceWin = win
    sessions.set(win.webContents.id, session)
    audiencePresenter.set(win.webContents.id, e.sender.id)
    const audienceWcId = win.webContents.id
    viewerWcIds.add(audienceWcId)
    win.once('ready-to-show', () => {
      win.show()
      fullScreenOnDisplay(win, external.bounds)
    })
    win.on('closed', () => {
      sessions.delete(audienceWcId)
      viewerWcIds.delete(audienceWcId)
      audiencePresenter.delete(audienceWcId)
      const s = presenterShows.get(e.sender.id)
      if (s?.audienceWin === win) s.audienceWin = null
    })
    if (runtime.rendererDevUrl) {
      const sep = runtime.rendererDevUrl.includes('?') ? '&' : '?'
      void win.loadURL(`${runtime.rendererDevUrl}${sep}mode=audience`)
    } else if (runtime.rendererFilePath) {
      void win.loadFile(runtime.rendererFilePath, { query: { mode: 'audience' } })
    }
    return { audience: true }
  })

  ipcMain.on('slides:presenter-sync', (e, state: ShowSyncState) => {
    const show = presenterShows.get(e.sender.id)
    if (!show) return
    show.lastSync = state
    const wc = show.audienceWin?.webContents
    if (wc && !wc.isDestroyed()) wc.send('slides:show-sync', state)
  })

  ipcMain.on('slides:presenter-ink', (e, ev: ShowInkEvent) => {
    const wc = presenterShows.get(e.sender.id)?.audienceWin?.webContents
    if (wc && !wc.isDestroyed()) wc.send('slides:show-ink', ev)
  })

  ipcMain.handle('slides:presenter-swap', (e) => {
    const show = presenterShows.get(e.sender.id)
    const aWin = show?.audienceWin
    const host = presenterHostWindow(e.sender)
    if (!show || !aWin || aWin.isDestroyed() || !host) return false
    const aDisplay = screen.getDisplayMatching(aWin.getBounds())
    const hDisplay = screen.getDisplayMatching(host.getBounds())
    if (aDisplay.id === hDisplay.id) return false
    fullScreenOnDisplay(aWin, hDisplay.bounds)
    // Move the presenter window to the audience's former screen, keeping its fullscreen form (HTML fullscreen uses native fullscreen)
    if (host.isFullScreen()) {
      host.once('leave-full-screen', () => {
        host.setBounds(aDisplay.workArea)
        host.setFullScreen(true)
      })
      host.setFullScreen(false)
    } else if (host.isSimpleFullScreen()) {
      host.setSimpleFullScreen(false)
      host.setBounds(aDisplay.workArea)
      host.setSimpleFullScreen(true)
    } else {
      host.setBounds(aDisplay.workArea)
    }
    return true
  })

  ipcMain.handle('slides:presenter-end', (e) => {
    closePresenterShow(e.sender.id)
  })

  ipcMain.handle('slides:audience-ready', (e) => {
    const pid = audiencePresenter.get(e.sender.id)
    return (pid != null ? presenterShows.get(pid)?.lastSync : null) ?? null
  })

  ipcMain.on('slides:audience-nav', (e, action: AudienceNavAction) => {
    const pid = audiencePresenter.get(e.sender.id)
    const wc = pid != null ? presenterShows.get(pid)?.presenterWc : null
    if (wc && !wc.isDestroyed()) wc.send('slides:audience-nav', action)
  })
}
