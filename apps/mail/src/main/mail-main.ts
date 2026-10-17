import { join } from 'node:path'
import { BrowserWindow, WebContentsView, app, dialog, ipcMain, shell } from 'electron'
import type { WebContents } from 'electron'
import {
  installContextMenu,
  installNavigationGuard,
  contextMenuLabels,
  installRendererProtocol,
  registerRendererScheme,
  rendererUrl,
  safeExternalUrl,
} from '@genoffice/electron-utils'
import { createI18n, getUiLang } from '@genoffice/i18n'
import { MAIL_CHANNELS } from '../shared/ipc'
import { isMailRef } from '../shared/mail-ref'

// View factory ONLY (Lane S owns). Compose IPC handlers live in a separate
// `compose-handlers.ts` (Lane C owns — never add them here). Read-path
// handlers (accounts/folders/messages/search) are registered by later lanes
// against the in-memory stub store (`./mail-store`).

const enClose = {
  closeUnsavedMsg: 'This mail tab has an unsent draft.',
  closeUnsavedDetail: 'Do you want to save the draft before closing?',
  btnSaveDraft: 'Save draft',
  btnDontSave: "Don't save",
  btnCancel: 'Cancel',
}

const tDlg = createI18n({
  zh: {
    closeUnsavedMsg: '此邮件标签页有未发送的草稿。',
    closeUnsavedDetail: '关闭前是否保存草稿？',
    btnSaveDraft: '保存草稿',
    btnDontSave: '不保存',
    btnCancel: '取消',
  },
  en: enClose,
  ja: enClose,
  ko: enClose,
  fr: enClose,
  de: enClose,
  es: enClose,
  th: enClose,
  id: enClose,
  ru: enClose,
  ar: enClose,
  pt: enClose,
  it: enClose,
  pl: enClose,
  cs: enClose,
  nl: enClose,
  ms: enClose,
  he: enClose,
  hi: enClose,
  'zh-TW': {
    closeUnsavedMsg: '此郵件標籤頁有未傳送的草稿。',
    closeUnsavedDetail: '關閉前是否儲存草稿？',
    btnSaveDraft: '儲存草稿',
    btnDontSave: '不儲存',
    btnCancel: '取消',
  },
  vi: enClose,
})
type DlgKey =
  'closeUnsavedMsg' | 'closeUnsavedDetail' | 'btnSaveDraft' | 'btnDontSave' | 'btnCancel'
const tm = (key: DlgKey): string => tDlg(getUiLang(), key)

interface RuntimePaths {
  preloadPath: string
  rendererUrl?: string
  rendererFile?: string
}

let runtime: RuntimePaths = { preloadPath: '' }

export function configureMailRuntime(paths: RuntimePaths): void {
  runtime = paths
}

/** mailRef behind each view (renderer state owns it; the tab record never holds `mail://`). */
const mailRefByWc = new Map<number, string>()
/** Unsaved-changes flags mirrored from the renderer; drives the draft prompt before closing a tab/window */
const dirtyByWc = new Set<number>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()

export function mailIsDirty(webContentsId: number): boolean {
  return dirtyByWc.has(webContentsId)
}

/** mailRef for the `mail:open-message` focus path; undefined for blank tabs. */
export function mailRefOf(webContentsId: number): string | undefined {
  return mailRefByWc.get(webContentsId)
}

/**
 * Close guard: true means proceed with closing. Clean → true; dirty →
 * Save draft / Discard draft / Cancel. On save, ask the renderer to persist
 * the draft and await the result; a failed save keeps the tab open.
 */
export async function requestMailClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  if (!dirtyByWc.has(contents.id) || contents.isDestroyed()) return true
  const options = {
    type: 'warning' as const,
    message: tm('closeUnsavedMsg'),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('btnSaveDraft'), tm('btnDontSave'), tm('btnCancel')],
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
    contents.send(MAIL_CHANNELS.closeSaveRequest)
  })
}

let ipcRegistered = false

function registerMailIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  ipcMain.on(MAIL_CHANNELS.dirtyChanged, (e, dirty: unknown) => {
    if (dirty === true) dirtyByWc.add(e.sender.id)
    else dirtyByWc.delete(e.sender.id)
  })

  ipcMain.on(MAIL_CHANNELS.closeSaveResult, (e, ok: unknown) => {
    const waiter = closeSaveWaiters.get(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // Language channel shared with other modules; removeHandler tolerates duplicate registration
  ipcMain.removeHandler(MAIL_CHANNELS.getLanguage)
  ipcMain.handle(MAIL_CHANNELS.getLanguage, () => getUiLang())
}

function grantAndTrack(wc: WebContents, mailRef?: string | null): void {
  const wcId = wc.id
  if (mailRef && isMailRef(mailRef)) mailRefByWc.set(wcId, mailRef)
  wc.setWindowOpenHandler(({ url }) => {
    // Per-call mailto allowlist at the mail call-site (Alya XR-1 record);
    // safeExternalUrl defaults stay http/https-only everywhere else.
    const target = safeExternalUrl(url, { allowedProtocols: ['http:', 'https:', 'mailto:'] })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    mailRefByWc.delete(wcId)
    dirtyByWc.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
  })
}

export function createMailView(mailRef?: string | null): WebContentsView {
  registerMailIpc()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  grantAndTrack(view.webContents, mailRef)
  void view.webContents.loadURL(rendererUrl(runtime.rendererUrl, 'mail'))
  return view
}

/** Standalone window mode: `npm run dev -w @genoffice/mail` */
export function startMailStandalone(): void {
  registerRendererScheme()
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  configureMailRuntime({
    preloadPath: join(__dirname, '../preload/index.js'),
    rendererUrl: process.env.ELECTRON_RENDERER_URL,
    rendererFile: join(__dirname, '../renderer/index.html'),
  })
  void app.whenReady().then(() => {
    installRendererProtocol({ mail: join(__dirname, '../renderer') })
    registerMailIpc()
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
    grantAndTrack(win.webContents, null)
    void win.loadURL(rendererUrl(runtime.rendererUrl, 'mail'))
  })
  app.on('window-all-closed', () => app.quit())
}
