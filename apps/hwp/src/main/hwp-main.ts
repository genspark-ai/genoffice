import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { BrowserWindow, WebContentsView, app, dialog, ipcMain, shell } from 'electron'
import type { WebContents } from 'electron'
import {
  configuredDefaultSaveDir,
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showSaveDialogWithMemory,
} from '@genoffice/electron-utils'
import { getUiLang } from '@genoffice/i18n'
import { HWP_CHANNELS } from '../shared/ipc'
import type { SaveHwpRequest, SaveHwpResult, SaveMode } from '../shared/ipc'
import { HWP_RE, bytesForSaveFormat, ensureHwpSavePath, saveFormatForPath } from '../shared/formats'
import { atomicWriteFile } from './atomic-write'
import { tm } from './dialogs'
import { startHwpLoopback } from './loopback'

interface RuntimePaths {
  preloadPath: string
  rendererUrl?: string
  rendererFile?: string
}

let runtime: RuntimePaths = { preloadPath: '' }
let loopbackOrigin: Promise<string> | null = null

export function configureHwpRuntime(paths: RuntimePaths): void {
  runtime = paths
}

async function rendererOrigin(): Promise<string> {
  if (runtime.rendererUrl) return runtime.rendererUrl
  if (!runtime.rendererFile) throw new Error('hwp: renderer path not configured')
  loopbackOrigin ??= startHwpLoopback(dirname(runtime.rendererFile))
  return loopbackOrigin
}

/** Open path per view, queued at tab creation; the renderer consumes it after mount. */
const openPathByWc = new Map<number, string>()
/** Current save target per view; absent = untitled document */
const savePathByWc = new Map<number, string>()
/** Unsaved-changes flags mirrored from the renderer */
const dirtyByWc = new Set<number>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
const saveWaiters = new Map<number, (ok: boolean) => void>()

let fileSavedHook: ((wc: WebContents, path: string) => void) | null = null

export function setHwpFileSavedHook(hook: (wc: WebContents, path: string) => void): void {
  fileSavedHook = hook
}

export function hwpIsDirty(webContentsId: number): boolean {
  return dirtyByWc.has(webContentsId)
}

export function hwpFilePath(webContentsId: number): string | undefined {
  return savePathByWc.get(webContentsId) ?? openPathByWc.get(webContentsId)
}

/** The file was renamed on disk — keep the queued/open path in sync. */
export function hwpFileRenamed(contents: WebContents, oldPath: string, newPath: string): void {
  const wcId = contents.id
  if (savePathByWc.get(wcId) === oldPath) savePathByWc.set(wcId, newPath)
  if (openPathByWc.get(wcId) === oldPath) openPathByWc.set(wcId, newPath)
  if (!contents.isDestroyed()) contents.send(HWP_CHANNELS.fileRenamed, newPath)
}

/**
 * Close guard: true means proceed with closing. Clean → true; dirty →
 * Save / Don't Save / Cancel. On Save, ask the renderer to export + write.
 */
export async function requestHwpClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  if (!dirtyByWc.has(contents.id) || contents.isDestroyed()) return true
  const options = {
    type: 'warning' as const,
    message: tm('closeUnsavedMsg'),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('btnSave'), tm('btnDontSave'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const { response } =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (response === 2) return false
  if (response === 1) return true
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(HWP_CHANNELS.closeSaveRequest)
  })
}

/** Menu Save / Save As: ask the renderer to export and save. */
export function requestHwpSave(contents: WebContents, mode: SaveMode): Promise<boolean> {
  if (contents.isDestroyed()) return Promise.resolve(false)
  if (mode === 'save' && !dirtyByWc.has(contents.id) && savePathByWc.has(contents.id)) {
    return Promise.resolve(true)
  }
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      saveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    saveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(HWP_CHANNELS.saveRequest, mode)
  })
}

function asNodeBuffer(raw: unknown): Buffer | null {
  if (Buffer.isBuffer(raw)) return raw
  if (raw instanceof Uint8Array) return Buffer.from(raw)
  if (raw instanceof ArrayBuffer) return Buffer.from(raw)
  if (ArrayBuffer.isView(raw)) {
    const view = raw as ArrayBufferView
    return Buffer.from(view.buffer, view.byteOffset, view.byteLength)
  }
  if (
    raw &&
    typeof raw === 'object' &&
    (raw as { type?: string }).type === 'Buffer' &&
    Array.isArray((raw as { data?: unknown }).data)
  ) {
    return Buffer.from((raw as { data: number[] }).data)
  }
  return null
}

async function resolveSaveTarget(
  e: Electron.IpcMainInvokeEvent,
  mode: SaveMode,
): Promise<string | null | 'canceled'> {
  const current = savePathByWc.get(e.sender.id)
  if (mode === 'save' && current) return current
  const win =
    BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
  const defaultPath = current
    ? join(dirname(current), basename(current))
    : join(configuredDefaultSaveDir(app), `${tm('untitledFile')}.hwp`)
  const picked = await showSaveDialogWithMemory(dialog, win, {
    title: tm('dlgSaveTitle'),
    defaultPath,
    filters: [
      { name: tm('filterHwp'), extensions: ['hwp'] },
      { name: tm('filterHwpx'), extensions: ['hwpx'] },
      { name: tm('filterHml'), extensions: ['hml'] },
    ],
  })
  if (picked.canceled || !picked.filePath) return 'canceled'
  return ensureHwpSavePath(picked.filePath)
}

let ipcRegistered = false

function registerHwpIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  ipcMain.handle(HWP_CHANNELS.consumePending, (e) => openPathByWc.get(e.sender.id) ?? null)

  ipcMain.handle(HWP_CHANNELS.readFile, async (e, path: unknown) => {
    if (typeof path !== 'string' || openPathByWc.get(e.sender.id) !== path) {
      throw new Error('hwp: path not granted to this view')
    }
    return await readFile(path)
  })

  ipcMain.handle(
    HWP_CHANNELS.save,
    async (e, request: SaveHwpRequest): Promise<SaveHwpResult> => {
      const waiter = saveWaiters.get(e.sender.id)
      saveWaiters.delete(e.sender.id)
      const done = (result: SaveHwpResult): SaveHwpResult => {
        waiter?.(result.ok && !('canceled' in result))
        return result
      }
      const hwp = asNodeBuffer(request?.hwp)
      const hwpx = asNodeBuffer(request?.hwpx)
      if (!hwp || !hwpx) return done({ ok: false, error: 'hwp: bad save request' })
      const hml = request?.hml !== undefined ? asNodeBuffer(request.hml) : undefined
      if (request?.hml !== undefined && !hml) {
        return done({ ok: false, error: 'hwp: bad HML bytes' })
      }
      const mode: SaveMode = request.mode === 'saveAs' ? 'saveAs' : 'save'
      try {
        const target = await resolveSaveTarget(e, mode)
        if (target === 'canceled') return done({ ok: true, canceled: true })
        if (!target) return done({ ok: false, error: 'hwp: no save target' })
        const format = saveFormatForPath(target)
        const bytes = bytesForSaveFormat(format, {
          hwp,
          hwpx,
          ...(hml ? { hml } : {}),
        })
        await atomicWriteFile(target, Buffer.from(bytes))
        const currentPath = savePathByWc.get(e.sender.id)
        savePathByWc.set(e.sender.id, target)
        openPathByWc.set(e.sender.id, target)
        dirtyByWc.delete(e.sender.id)
        if (currentPath !== target) fileSavedHook?.(e.sender, target)
        return done({ ok: true, path: target })
      } catch (err) {
        return done({ ok: false, error: err instanceof Error ? err.message : String(err) })
      }
    },
  )

  ipcMain.on(HWP_CHANNELS.dirtyChanged, (e, dirty: unknown) => {
    if (dirty === true) dirtyByWc.add(e.sender.id)
    else dirtyByWc.delete(e.sender.id)
  })

  ipcMain.on(HWP_CHANNELS.closeSaveResult, (e, ok: unknown) => {
    const waiter = closeSaveWaiters.get(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  ipcMain.on(HWP_CHANNELS.saveRequestAck, (e, ok: unknown) => {
    const waiter = saveWaiters.get(e.sender.id)
    saveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  ipcMain.removeHandler(HWP_CHANNELS.getLanguage)
  ipcMain.handle(HWP_CHANNELS.getLanguage, () => getUiLang())
}

function grantAndTrack(wc: WebContents, openPath?: string | null): void {
  const wcId = wc.id
  if (openPath && existsSync(openPath) && HWP_RE.test(openPath)) {
    openPathByWc.set(wcId, openPath)
    savePathByWc.set(wcId, openPath)
  }
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: ['http:', 'https:', 'mailto:'] })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    openPathByWc.delete(wcId)
    savePathByWc.delete(wcId)
    dirtyByWc.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
    saveWaiters.get(wcId)?.(false)
    saveWaiters.delete(wcId)
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
  void rendererOrigin().then((url) => view.webContents.loadURL(url))
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
  void app.whenReady().then(async () => {
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
    await win.loadURL(await rendererOrigin())
  })
  app.on('window-all-closed', () => app.quit())
}
