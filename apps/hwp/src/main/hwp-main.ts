import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { BrowserWindow, WebContentsView, app, ipcMain, shell } from 'electron'
import type { WebContents } from 'electron'
import {
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
} from '@genoffice/electron-utils'
import { getUiLang } from '@genoffice/i18n'
import { HWP_CHANNELS } from '../shared/ipc'
import { HWP_RE } from '../shared/formats'

interface RuntimePaths {
  preloadPath: string
  rendererUrl?: string
  rendererFile?: string
}

let runtime: RuntimePaths = { preloadPath: '' }

export function configureHwpRuntime(paths: RuntimePaths): void {
  runtime = paths
}

/** Open path per view, queued at tab creation; the renderer consumes it after mount. */
const openPathByWc = new Map<number, string>()

export function hwpIsDirty(_webContentsId: number): boolean {
  return false
}

export function hwpFilePath(webContentsId: number): string | undefined {
  return openPathByWc.get(webContentsId)
}

/** The file was renamed on disk — keep the queued/open path in sync. */
export function hwpFileRenamed(contents: WebContents, oldPath: string, newPath: string): void {
  const wcId = contents.id
  if (openPathByWc.get(wcId) === oldPath) openPathByWc.set(wcId, newPath)
  if (!contents.isDestroyed()) contents.send(HWP_CHANNELS.fileRenamed, newPath)
}

/** No unsaved edits in this slice — close always proceeds. */
export async function requestHwpClose(
  _contents: WebContents,
  _parent?: BrowserWindow | null,
): Promise<boolean> {
  return true
}

let ipcRegistered = false

function registerHwpIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  ipcMain.handle(HWP_CHANNELS.consumePending, (e) => openPathByWc.get(e.sender.id) ?? null)

  ipcMain.removeHandler(HWP_CHANNELS.getLanguage)
  ipcMain.handle(HWP_CHANNELS.getLanguage, () => getUiLang())
}

function grantAndTrack(wc: WebContents, openPath?: string | null): void {
  const wcId = wc.id
  if (openPath && existsSync(openPath) && HWP_RE.test(openPath)) {
    openPathByWc.set(wcId, openPath)
  }
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: ['http:', 'https:', 'mailto:'] })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    openPathByWc.delete(wcId)
  })
}

export function createHwpView(openPath?: string | null): WebContentsView {
  registerHwpIpc()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  grantAndTrack(view.webContents, openPath)
  if (runtime.rendererUrl) void view.webContents.loadURL(runtime.rendererUrl)
  else if (runtime.rendererFile) void view.webContents.loadFile(runtime.rendererFile)
  return view
}

/** Standalone window mode: `npm run dev -w @genoffice/hwp` */
export function startHwpStandalone(): void {
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  configureHwpRuntime({
    preloadPath: join(__dirname, '../preload/index.js'),
    rendererUrl: process.env.ELECTRON_RENDERER_URL,
    rendererFile: join(__dirname, '../renderer/index.html'),
  })
  void app.whenReady().then(() => {
    registerHwpIpc()
    const win = new BrowserWindow({
      width: 1200,
      height: 850,
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
      },
    })
    const argPath = process.argv.slice(1).find((a) => HWP_RE.test(a) && existsSync(a))
    grantAndTrack(win.webContents, argPath)
    if (runtime.rendererUrl) void win.loadURL(runtime.rendererUrl)
    else if (runtime.rendererFile) void win.loadFile(runtime.rendererFile)
  })
  app.on('window-all-closed', () => app.quit())
}
