import { basename } from 'node:path'
import { BrowserWindow } from 'electron'
import type { WebContents, WebContentsView } from 'electron'

import {
  docsQueryDirty,
  requestDocsClose,
  setActiveDocsResolver,
  teardownDocsRenderer,
} from '../../../docs/src/main/docs-main'
import {
  requestSheetsClose,
  setActiveSheetsWebContents,
  sheetsPendingEditCount,
} from '../../../sheets/src/main/sheets-main'
import { canonicalPath } from './tab-manager'
import type { OpenDocumentTab, TabKind } from '../shared/tabs-api'

/**
 * Detached editor windows ("Open in New Window" on a docs/sheets tab): the
 * tab's live WebContentsView is reparented into its own BrowserWindow, so the
 * document — including unsaved edits — moves without a reload. The window uses
 * the native frame (the shell's tab strip is the drag surface in tab mode;
 * a detached window has none), and the focused window claims the
 * process-global menu/active-editor targets — the shell's tab manager takes
 * them back on its own window's focus (refreshActiveTargets).
 */
interface DetachedRecord {
  window: BrowserWindow
  view: WebContentsView
  kind: TabKind
  filePath?: string
  /** tear the window down with no save prompt (agent/MCP close) */
  closeWithoutPrompt: () => void
}

/** keyed by the view's webContents id — the same key the app modules use */
const detached = new Map<number, DetachedRecord>()

/** tab-style ids for the control host / MCP: `detached:<webContents id>` */
const ID_PREFIX = 'detached:'

let onChanged: () => void = () => {}

/** fires whenever a detached window opens, closes, or changes file (open-documents publishing) */
export function setDetachedChangedListener(listener: () => void): void {
  onChanged = listener
}

function recordById(id: string): DetachedRecord | undefined {
  if (!id.startsWith(ID_PREFIX)) return undefined
  const rec = detached.get(Number(id.slice(ID_PREFIX.length)))
  return rec && !rec.window.isDestroyed() ? rec : undefined
}

function bringToFront(win: BrowserWindow): void {
  if (win.isMinimized()) win.restore()
  win.show()
  win.focus()
}

const DEFAULT_SIZE: Record<string, { width: number; height: number }> = {
  docs: { width: 1360, height: 900 },
  sheets: { width: 1440, height: 900 },
}

export function isDetachedEditorWindow(win: BrowserWindow): boolean {
  for (const rec of detached.values()) if (rec.window === win) return true
  return false
}

/** focus the detached window editing this file, for open-path dedup (Home, recents, CLI) */
export function focusDetachedByPath(path: string): boolean {
  const wanted = canonicalPath(path)
  for (const rec of detached.values()) {
    if (rec.filePath && canonicalPath(rec.filePath) === wanted && !rec.window.isDestroyed()) {
      bringToFront(rec.window)
      return true
    }
  }
  return false
}

/** control-host lookup: a detached editor answers like a tab, with a `detached:` id */
export function findDetachedTabByPath(
  path: string,
): { id: string; kind: TabKind; webContents: WebContents } | undefined {
  const wanted = canonicalPath(path)
  for (const [wcId, rec] of detached) {
    if (rec.filePath && canonicalPath(rec.filePath) === wanted && !rec.window.isDestroyed()) {
      return { id: ID_PREFIX + wcId, kind: rec.kind, webContents: rec.view.webContents }
    }
  }
  return undefined
}

/** the detached window hosting this webContents (dialog parenting: fromWebContents
 *  cannot resolve a WebContentsView's host) */
export function detachedWindowForWebContents(wcId: number): BrowserWindow | undefined {
  const rec = detached.get(wcId)
  return rec && !rec.window.isDestroyed() ? rec.window : undefined
}

/** editor kind of the focused detached window, if a detached window has focus */
export function focusedDetachedKind(): TabKind | undefined {
  for (const rec of detached.values()) {
    if (!rec.window.isDestroyed() && rec.window.isFocused()) return rec.kind
  }
  return undefined
}

export function isDetachedTabId(id: string): boolean {
  return recordById(id) !== undefined
}

/** bring a detached editor forward; false when the id is not a detached window */
export function activateDetached(id: string): boolean {
  const rec = recordById(id)
  if (!rec) return false
  bringToFront(rec.window)
  return true
}

export function detachedWebContentsFor(id: string): WebContents | undefined {
  return recordById(id)?.view.webContents
}

export function closeDetachedWithoutPrompt(id: string): boolean {
  const rec = recordById(id)
  if (!rec) return false
  rec.closeWithoutPrompt()
  return true
}

/** the detached editors in the same shape as TabManager.openDocuments (agent/MCP targets) */
export async function detachedOpenDocuments(): Promise<OpenDocumentTab[]> {
  const out: OpenDocumentTab[] = []
  for (const [wcId, rec] of detached) {
    if (rec.window.isDestroyed()) continue
    const wc = rec.view.webContents
    const dirty =
      rec.kind === 'sheets'
        ? sheetsPendingEditCount(wcId) > 0
        : rec.kind === 'docs'
          ? await docsQueryDirty(wc)
          : false
    out.push({
      id: ID_PREFIX + wcId,
      kind: rec.kind as OpenDocumentTab['kind'],
      title: rec.window.getTitle(),
      ...(rec.filePath ? { filePath: rec.filePath } : {}),
      active: rec.window.isFocused(),
      dirty,
    })
  }
  return out
}

/** Save As / open-in-place landed on a new path — keep the window title and
 *  the dedup registry in sync (same contract as TabManager.setTabFileFor). */
export function detachedSetFileFor(webContentsId: number, filePath: string): void {
  const rec = detached.get(webContentsId)
  if (!rec) return
  rec.filePath = filePath
  if (!rec.window.isDestroyed()) rec.window.setTitle(basename(filePath))
  onChanged()
}

/** a rename on disk (Home list) — follow it, same contract as renameTabFile */
export function detachedRenameFile(
  oldPath: string,
  newPath: string,
): { kind: TabKind; webContents: WebContents } | undefined {
  const wanted = canonicalPath(oldPath)
  for (const rec of detached.values()) {
    if (rec.filePath && canonicalPath(rec.filePath) === wanted) {
      rec.filePath = newPath
      if (!rec.window.isDestroyed()) rec.window.setTitle(basename(newPath))
      onChanged()
      return { kind: rec.kind, webContents: rec.view.webContents }
    }
  }
  return undefined
}

/** every detached window's open file (agent/MCP open-documents publishing) */
export function detachedFilePaths(): string[] {
  const paths: string[] = []
  for (const rec of detached.values())
    if (rec.filePath && !rec.window.isDestroyed()) paths.push(rec.filePath)
  return paths
}

export function createDetachedEditorWindow(options: {
  view: WebContentsView
  kind: TabKind
  title: string
  filePath?: string
  applyMenuFor: (kind: TabKind) => void
}): BrowserWindow {
  const { view, kind, title, filePath, applyMenuFor } = options
  const size = DEFAULT_SIZE[kind] ?? { width: 1360, height: 900 }
  const win = new BrowserWindow({
    ...size,
    minWidth: 720,
    minHeight: 550,
    title,
    // native frame: the editor view fills the window, so the OS title bar is
    // the only drag surface / window-controls host
    autoHideMenuBar: true,
  })
  const wcId = view.webContents.id
  let closeConfirmed = false
  // Detach the view before destroying the window so the docs teardown path
  // (which must never close the webContents — see TabManager.closeTab) stays
  // in control of the renderer's lifetime.
  const tearDown = () => {
    if (win.isDestroyed()) return
    closeConfirmed = true
    win.contentView.removeChildView(view)
    detached.delete(wcId)
    if (kind === 'docs') teardownDocsRenderer(view.webContents)
    else view.webContents.close()
    win.destroy()
  }
  detached.set(wcId, { window: win, view, kind, filePath, closeWithoutPrompt: tearDown })
  onChanged()

  win.contentView.addChildView(view)
  const layout = () => {
    if (win.isDestroyed()) return
    const bounds = win.getContentBounds()
    view.setBounds({ x: 0, y: 0, width: bounds.width, height: bounds.height })
  }
  // same Linux/X11 quirk as the shell window: `resize` fires before the WM
  // applies the new size, so a follow-up layout on the next tick is required
  win.on('resize', () => {
    layout()
    setImmediate(() => layout())
  })
  layout()
  view.setVisible(true)
  view.webContents.focus()

  // the title is owned here (detach-time tab title, Save As updates);
  // the renderer's static <title> must not overwrite it
  win.on('page-title-updated', (event) => event.preventDefault())

  // The focused window owns the process-global menu-command targets; the
  // shell's tab manager re-claims both on its own window's focus.
  win.on('focus', () => {
    if (kind === 'docs')
      setActiveDocsResolver(() => (view.webContents.isDestroyed() ? null : view.webContents))
    if (kind === 'sheets') setActiveSheetsWebContents(view.webContents)
    applyMenuFor(kind)
  })

  // Unsaved-changes guard, same helpers as the tab close path.
  win.on('close', (event) => {
    if (closeConfirmed) return
    event.preventDefault()
    void (async () => {
      let proceed = true
      if (kind === 'sheets' && sheetsPendingEditCount(wcId) > 0) {
        proceed = await requestSheetsClose(view.webContents, win)
      } else if (kind === 'docs' && (await docsQueryDirty(view.webContents))) {
        proceed = await requestDocsClose(view.webContents, win)
      }
      if (proceed) tearDown()
    })()
  })
  win.on('closed', () => {
    detached.delete(wcId)
    onChanged()
  })
  return win
}
