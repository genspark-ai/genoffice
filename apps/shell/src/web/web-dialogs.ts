import * as electronDialog from 'electron'
import type { MessageBoxOptions, OpenDialogOptions, SaveDialogOptions } from 'electron'
import { readdir } from 'node:fs/promises'
import { basename, dirname, extname, join, resolve, sep } from 'node:path'
import { existsSync, statSync } from 'node:fs'
import type { WebTabContents } from './web-contents'

/**
 * Native open/save/message dialogs belong to the server machine; for requests
 * arriving from a browser tab they are served in the tab itself. The patch is
 * inert outside a web dispatch (desktop windows keep the native dialogs), and
 * inside a web dispatch every dialog call resolves through that tab's bridge —
 * the handler code across all apps keeps calling the same Electron APIs.
 */

export type WebDialogOptions = {
  kind: 'save' | 'open' | 'message'
  title?: string
  defaultPath?: string
  buttonLabel?: string
  filters?: { name: string; extensions: string[] }[]
  properties?: string[]
  message?: string
  detail?: string
  buttons?: string[]
}

export type WebDialogResult =
  | { canceled: true }
  | { canceled: false; filePath: string }
  | { canceled: false; filePaths: string[] }
  | { canceled: false; response: number; checkboxChecked?: boolean }

const pending = new Map<number, (result: WebDialogResult) => void>()
let nextDialogId = 1

/** deliver a dialog request to one browser tab and wait for its overlay */
export function serveWebDialog(
  tab: WebTabContents,
  options: WebDialogOptions,
): Promise<WebDialogResult> {
  const id = nextDialogId++
  dialogOwners.set(id, tab)
  return new Promise<WebDialogResult>((resolvePromise) => {
    pending.set(id, resolvePromise)
    tab.deliver('web:dialog', [{ id, options }])
  }).finally(() => {
    pending.delete(id)
    dialogOwners.delete(id)
  })
}

/** bridge callback: the tab answered a dialog request */
export function resolveWebDialog(id: number, payload: WebDialogResult): void {
  pending.get(id)?.(payload)
}

export function cancelWebDialogsFor(tab: WebTabContents): void {
  // resolvers keyed by id are tab-agnostic; a closed tab's overlay can never
  // answer, so drop everything still pending from that tab's dispatch. The
  // dispatch wrapper below scopes serving, so any pending dialog belongs to
  // the disconnecting tab.
  for (const [id, resolver] of [...pending]) {
    if (dialogOwners.get(id) === tab) {
      pending.delete(id)
      resolver({ canceled: true })
    }
  }
}

const dialogOwners = new Map<number, WebTabContents>()

/** the tab whose handler dispatch is running; null during desktop dispatches */
let dispatchTab: WebTabContents | null = null

export function beginWebDispatch(tab: WebTabContents): void {
  dispatchTab = tab
}

export function endWebDispatch(): void {
  dispatchTab = null
}

function isWebDispatch(): boolean {
  return dispatchTab !== null
}

type ShowSave = (options: SaveDialogOptions) => Promise<{ canceled: boolean; filePath?: string }>
type ShowOpen = (options: OpenDialogOptions) => Promise<{ canceled: boolean; filePaths: string[] }>
type ShowMessage = (options: MessageBoxOptions) => Promise<{ response: number }>

/** Install once at web-server startup; wraps the Electron dialog methods. */
export function installWebDialogPatch(): void {
  const d = electronDialog.dialog as unknown as {
    showSaveDialog: ShowSave
    showOpenDialog: ShowOpen
    showMessageBox: ShowMessage
  }
  const nativeSave = d.showSaveDialog.bind(d)
  const nativeOpen = d.showOpenDialog.bind(d)
  const nativeMessage = d.showMessageBox.bind(d)

  d.showSaveDialog = (options: SaveDialogOptions) => {
    if (!isWebDispatch()) return nativeSave(options)
    return serveWebDialog(dispatchTab!, saveOptionsOf(options)).then((result) =>
      result.canceled || !('filePath' in result)
        ? { canceled: true }
        : { canceled: false, filePath: result.filePath },
    )
  }
  d.showOpenDialog = (options: OpenDialogOptions) => {
    if (!isWebDispatch()) return nativeOpen(options)
    return serveWebDialog(dispatchTab!, {
      kind: 'open',
      title: options.title,
      defaultPath: Array.isArray(options.defaultPath) ? undefined : options.defaultPath,
      buttonLabel: options.buttonLabel,
      filters: options.filters,
      properties: options.properties,
    }).then((result) =>
      result.canceled || !('filePaths' in result)
        ? { canceled: true, filePaths: [] }
        : { canceled: false, filePaths: result.filePaths },
    )
  }
  d.showMessageBox = (options: MessageBoxOptions) => {
    if (!isWebDispatch()) return nativeMessage(options)
    return serveWebDialog(dispatchTab!, {
      kind: 'message',
      title: options.title,
      message: options.message,
      detail: options.detail,
      buttons: options.buttons,
    }).then((result) =>
      result.canceled || !('response' in result)
        ? { response: 0 }
        : { response: result.response, checkboxChecked: result.checkboxChecked },
    ) as Promise<{ response: number }>
  }
}

function saveOptionsOf(options: SaveDialogOptions): WebDialogOptions {
  return {
    kind: 'save',
    title: options.title,
    defaultPath: options.defaultPath,
    buttonLabel: options.buttonLabel,
    filters: options.filters,
  }
}

// ---- backing channels the browser overlay calls through the bridge ----

export interface DirEntry {
  name: string
  dir: boolean
}

/** entries of one directory for the overlay's listing; dirs first, extension
 *  filtering only applies to files, and unreadable dirs come back empty */
export async function listWebDialogDir(dir: string, exts: string[]): Promise<DirEntry[]> {
  let target = resolve(dir)
  if (!existsSync(target)) return []
  if (!statSync(target).isDirectory()) target = dirname(target)
  let names: string[]
  try {
    names = await readdir(target)
  } catch {
    return []
  }
  const lower = new Set(exts.map((e) => e.toLowerCase()))
  const entries: DirEntry[] = []
  for (const name of names.sort()) {
    if (name.startsWith('.')) continue
    let isDir: boolean
    try {
      isDir = statSync(join(target, name)).isDirectory()
    } catch {
      continue
    }
    if (!isDir && lower.size > 0 && !lower.has(extname(name).toLowerCase().replace(/^\./, ''))) {
      continue
    }
    entries.push({ name, dir: isDir })
  }
  entries.sort((a, b) => Number(b.dir) - Number(a.dir) || a.name.localeCompare(b.name))
  return entries
}

/** the directories the overlay can start from: the dialog's own default first */
export function webDialogStartDirs(defaultPath?: string): string[] {
  const starts: string[] = []
  if (defaultPath) {
    const p = resolve(defaultPath)
    starts.push(existsSync(p) && statSync(p).isDirectory() ? p : dirname(p))
  }
  return starts.length > 0 ? starts : [resolve(sep === '/' ? '/' : 'C:\\')]
}

/** final validation before a dialog response reaches an app handler */
export function validateWebDialogPath(path: string, kind: 'save' | 'open'): string | null {
  if (typeof path !== 'string' || path.length === 0 || path.includes('\0')) return null
  const p = resolve(path)
  if (kind === 'open') return existsSync(p) && statSync(p).isFile() ? p : null
  const dir = dirname(p)
  return existsSync(dir) && statSync(dir).isDirectory() ? p : null
}

export function webDialogFileName(path: string): string {
  return basename(path)
}
