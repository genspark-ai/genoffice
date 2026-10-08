import { basename } from 'node:path'
import { BrowserWindow, screen } from 'electron'
import type { Rectangle, WebContents, WebContentsView } from 'electron'

import {
  docsQueryDirty,
  requestDocsClose,
  setActiveDocsResolver,
  teardownDocsRenderer,
} from '../../../docs/src/main/docs-main'
import {
  requestSheetsClose,
  resetSheetsShuttingDown,
  setActiveSheetsWebContents,
  sheetsPendingEditCount,
} from '../../../sheets/src/main/sheets-main'
import {
  requestSlidesClose,
  setActiveSlidesWebContents,
  slidesIsDirty,
} from '../../../slides/src/main/slides-main'
import { pdfIsDirty, requestPdfClose } from '../../../pdf/src/main/pdf-main'
import { markdownIsDirty, requestMarkdownClose } from '../../../markdown/src/main/markdown-main'
import { htmlIsDirty, requestHtmlClose } from '../../../html/src/main/html-main'
import { canonicalPath } from './tab-manager'
import type { DetachedTab } from './tab-manager'
import type { OpenDocumentTab, TabKind } from '../shared/tabs-api'
import { DOCK_DWELL_MS, dockBand, pointInRect } from '../shared/tab-drag-geometry'

/**
 * Detached editor windows: a tab's live WebContentsView reparented into its
 * own BrowserWindow, so the document — including unsaved edits — moves
 * without a reload. Three ways in, two ways out:
 *
 *   in:  "Open in New Window" on a tab; dragging a tab off the strip
 *        (tear-off: the new window follows the still-held pointer);
 *   out: closing the window (own unsaved-changes guard); dragging the
 *        window's title bar over the shell's tab strip, which docks the
 *        document back as a tab (also mid tear-off, when the held pointer
 *        returns to the strip).
 *
 * The window uses the native frame (the shell's tab strip is the drag surface
 * in tab mode; a detached window has none), and the focused window claims the
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
  /** the view left for the shell strip: the pending destroy must not touch it */
  released: boolean
  /** cursor currently inside the shell strip's dock band during a native drag */
  hovering: boolean
  /** dwell timer: docks once the cursor has rested over the strip */
  dockTimer: ReturnType<typeof setTimeout> | null
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

/**
 * The shell side of docking, registered by the shell window: where its strip
 * is on screen, how to show the insertion indicator, and how to take a
 * document back. Null while there is no shell window to dock into.
 */
export interface DockHost {
  /** screen rect the cursor must hover to dock; null when the shell is gone or hidden */
  band: () => Rectangle | null
  /** show (window-CSS x) or clear (null) the strip's insertion indicator */
  preview: (x: number | null) => void
  /** put the document back into the strip (at the slot the strip last reported) */
  dock: (tab: DetachedTab) => void
}
let dockHost: DockHost | null = null
export function setDockHost(host: DockHost | null): void {
  dockHost = host
}

/** the one window mid tear-off: created under the held pointer, following it */
let torn: { rec: DetachedRecord; grabDx: number; grabDy: number } | null = null

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
  slides: { width: 1440, height: 900 },
  pdf: { width: 1200, height: 900 },
  markdown: { width: 1200, height: 850 },
  html: { width: 1200, height: 850 },
}

/** where the pointer sits inside a freshly torn-off window's title bar */
const TEAR_GRAB = { dx: 120, dy: 14 }

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
  return focusedDetachedTab()?.kind
}

/** the focused detached window's document, in the shape the menu targets use
 *  (pdf / markdown / html menus resolve their target through the shell) */
export function focusedDetachedTab():
  | {
      id: string
      kind: TabKind
      webContents: WebContents
      filePath?: string
      window: BrowserWindow
    }
  | undefined {
  for (const [wcId, rec] of detached) {
    if (!rec.window.isDestroyed() && rec.window.isFocused()) {
      return {
        id: ID_PREFIX + wcId,
        kind: rec.kind,
        webContents: rec.view.webContents,
        ...(rec.filePath ? { filePath: rec.filePath } : {}),
        window: rec.window,
      }
    }
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

/** unsaved-changes state of one detached document, whichever family owns it */
async function detachedIsDirty(rec: DetachedRecord): Promise<boolean> {
  const wc = rec.view.webContents
  if (wc.isDestroyed()) return false
  switch (rec.kind) {
    case 'sheets':
      return sheetsPendingEditCount(wc.id) > 0
    case 'docs':
      return docsQueryDirty(wc)
    case 'slides':
      return slidesIsDirty(wc.id)
    case 'pdf':
      return pdfIsDirty(wc.id)
    case 'markdown':
      return markdownIsDirty(wc.id)
    case 'html':
      return htmlIsDirty(wc.id)
    default:
      return false
  }
}

/** run the family's save/don't-save/cancel prompt when the document is dirty;
 *  true when the window may close (same helpers as the tab close path) */
async function confirmDetachedClose(rec: DetachedRecord): Promise<boolean> {
  if (!(await detachedIsDirty(rec))) return true
  const wc = rec.view.webContents
  const win = rec.window
  switch (rec.kind) {
    case 'sheets':
      return requestSheetsClose(wc, win)
    case 'docs':
      return requestDocsClose(wc, win)
    case 'slides':
      return requestSlidesClose(wc, win)
    case 'pdf':
      return requestPdfClose(wc, win)
    case 'markdown':
      return requestMarkdownClose(wc, win)
    case 'html':
      return requestHtmlClose(wc, win)
    default:
      return true
  }
}

/** the detached editors in the same shape as TabManager.openDocuments (agent/MCP targets) */
export async function detachedOpenDocuments(): Promise<OpenDocumentTab[]> {
  const out: OpenDocumentTab[] = []
  for (const [wcId, rec] of detached) {
    if (rec.window.isDestroyed()) continue
    const dirty = await detachedIsDirty(rec)
    // the dirty query yielded: the window may have closed meanwhile
    if (rec.window.isDestroyed()) continue
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

/** an untitled document named itself before its first save (html: from the first AI request) */
export function detachedSetTitleFor(webContentsId: number, title: string): void {
  const rec = detached.get(webContentsId)
  if (!rec || rec.filePath || rec.window.isDestroyed()) return
  rec.window.setTitle(title)
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

function clearDockTimer(rec: DetachedRecord): void {
  if (rec.dockTimer) clearTimeout(rec.dockTimer)
  rec.dockTimer = null
}

function clearDockHover(rec: DetachedRecord): void {
  clearDockTimer(rec)
  if (!rec.hovering) return
  rec.hovering = false
  dockHost?.preview(null)
}

/**
 * Lift the document out of its window without closing the renderer: the view
 * is unparented and the window destroyed (no close prompt — nothing is being
 * discarded). The caller re-homes the view.
 */
function releaseDetached(rec: DetachedRecord): DetachedTab {
  const { window: win, view, kind, filePath } = rec
  clearDockHover(rec)
  rec.released = true
  if (torn?.rec === rec) torn = null
  const title = win.isDestroyed() ? (filePath ? basename(filePath) : '') : win.getTitle()
  detached.delete(view.webContents.id)
  if (!win.isDestroyed()) {
    win.contentView.removeChildView(view)
    win.destroy()
  }
  onChanged()
  return { view, kind, title, filePath }
}

function dockDetachedWindow(rec: DetachedRecord): void {
  const host = dockHost
  if (!host || rec.released || rec.window.isDestroyed()) return
  host.dock(releaseDetached(rec))
}

/**
 * Native title-bar drag: sample the cursor against the shell strip's dock
 * band. Inside → indicator + dwell timer (the OS gives no release event on
 * every platform, so resting over the strip is the commit, Chrome-style);
 * outside → clear both.
 */
function trackDockHover(rec: DetachedRecord): void {
  if (rec.released || !dockHost) return
  const band = dockHost.band()
  const cursor = screen.getCursorScreenPoint()
  if (!band || !pointInRect(cursor.x, cursor.y, band)) {
    clearDockHover(rec)
    return
  }
  rec.hovering = true
  dockHost.preview(cursor.x - band.x)
  clearDockTimer(rec)
  rec.dockTimer = setTimeout(() => {
    rec.dockTimer = null
    const nowBand = dockHost?.band()
    const now = screen.getCursorScreenPoint()
    if (nowBand && pointInRect(now.x, now.y, nowBand)) dockDetachedWindow(rec)
    else clearDockHover(rec)
  }, DOCK_DWELL_MS)
}

/** screen rect of the shell's dock band for a given shell window (null when hidden) */
export function dockBandFor(shellWindow: BrowserWindow | null): Rectangle | null {
  if (!shellWindow || shellWindow.isDestroyed() || !shellWindow.isVisible()) return null
  if (shellWindow.isMinimized()) return null
  return dockBand(shellWindow.getContentBounds())
}

/** frame origin that puts the pointer (at) inside the new window's title bar,
 *  kept on the pointer's display */
function tearOffOrigin(at: { x: number; y: number }, size: { width: number; height: number }) {
  const area = screen.getDisplayNearestPoint(at).workArea
  const x = Math.min(Math.max(at.x - TEAR_GRAB.dx, area.x), area.x + area.width - size.width)
  const y = Math.min(Math.max(at.y - TEAR_GRAB.dy, area.y), area.y + area.height - size.height)
  return { x: Math.round(x), y: Math.round(y) }
}

export function createDetachedEditorWindow(options: {
  view: WebContentsView
  kind: TabKind
  title: string
  filePath?: string
  applyMenuFor: (kind: TabKind) => void
  /**
   * tear-off: place the window so the pointer at this screen point sits in
   * its title bar, and show it without focus — the shell keeps receiving the
   * held pointer's moves and steers the window until release
   */
  tearOffAt?: { x: number; y: number }
}): BrowserWindow {
  const { view, kind, title, filePath, applyMenuFor, tearOffAt } = options
  const size = DEFAULT_SIZE[kind] ?? { width: 1360, height: 900 }
  const origin = tearOffAt ? tearOffOrigin(tearOffAt, size) : {}
  const win = new BrowserWindow({
    ...size,
    ...origin,
    minWidth: 720,
    minHeight: 550,
    title,
    show: !tearOffAt,
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
    if (win.isDestroyed() || rec.released) return
    closeConfirmed = true
    clearDockHover(rec)
    if (torn?.rec === rec) torn = null
    win.contentView.removeChildView(view)
    detached.delete(wcId)
    if (kind === 'docs') teardownDocsRenderer(view.webContents)
    else view.webContents.close()
    win.destroy()
  }
  const rec: DetachedRecord = {
    window: win,
    view,
    kind,
    filePath,
    closeWithoutPrompt: tearDown,
    released: false,
    hovering: false,
    dockTimer: null,
  }
  detached.set(wcId, rec)
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
  if (tearOffAt) {
    const pos = win.getPosition()
    torn = { rec, grabDx: tearOffAt.x - pos[0], grabDy: tearOffAt.y - pos[1] }
    win.showInactive()
  } else {
    view.webContents.focus()
  }

  // the title is owned here (detach-time tab title, Save As updates);
  // the renderer's static <title> must not overwrite it
  win.on('page-title-updated', (event) => event.preventDefault())

  // The focused window owns the process-global menu-command targets; the
  // shell's tab manager re-claims both on its own window's focus.
  win.on('focus', () => {
    if (kind === 'docs')
      setActiveDocsResolver(() => (view.webContents.isDestroyed() ? null : view.webContents))
    if (kind === 'sheets') setActiveSheetsWebContents(view.webContents)
    if (kind === 'slides') setActiveSlidesWebContents(view.webContents)
    applyMenuFor(kind)
  })

  // Dragging the window by its title bar over the shell's tab strip docks the
  // document back. A torn-off window is steered by the shell's held pointer
  // instead, and that path decides about docking itself.
  win.on('move', () => {
    if (torn?.rec === rec) return
    trackDockHover(rec)
  })
  // Windows reports the end of a native drag as `moved`: commit right away
  // instead of waiting out the dwell. (macOS fires `moved` on every step, so
  // it is no release signal there.)
  if (process.platform === 'win32') {
    win.on('moved', () => {
      if (torn?.rec === rec || !rec.hovering) return
      trackDockHover(rec)
      if (rec.hovering) dockDetachedWindow(rec)
    })
  }

  // Unsaved-changes guard, same helpers as the tab close path.
  // a second close event while the prompt is open must not re-enter the
  // prompt (the same double-trigger the main window guards against)
  let promptInFlight = false
  win.on('close', (event) => {
    if (closeConfirmed || rec.released) return
    event.preventDefault()
    if (promptInFlight) return
    promptInFlight = true
    void (async () => {
      try {
        // a denied close vetoes any quit that was in flight: the sheets close
        // guard must prompt again on later closes instead of silently proceeding
        if (await confirmDetachedClose(rec)) tearDown()
        else resetSheetsShuttingDown()
      } catch (err) {
        console.error('[shell] detached window close prompt failed:', err)
        resetSheetsShuttingDown()
      } finally {
        promptInFlight = false
      }
    })()
  })
  win.on('closed', () => {
    clearDockHover(rec)
    if (torn?.rec === rec) torn = null
    detached.delete(wcId)
    onChanged()
  })
  return win
}

/** steer the torn-off window after the held pointer (screen coordinates) */
export function dragTornWindow(x: number, y: number): void {
  if (!torn || torn.rec.window.isDestroyed()) return
  torn.rec.window.setPosition(Math.round(x - torn.grabDx), Math.round(y - torn.grabDy))
}

/** the pointer was released: the torn-off window is an ordinary detached window now */
export function endTornDrag(): void {
  const rec = torn?.rec
  torn = null
  if (!rec || rec.window.isDestroyed()) return
  rec.window.focus()
  rec.view.webContents.focus()
}

/** the held pointer came back over the strip: take the document out of the
 *  torn-off window (destroying it) so the shell can make it a tab again */
export function takeTornTab(): DetachedTab | null {
  const rec = torn?.rec
  torn = null
  if (!rec || rec.released || rec.window.isDestroyed()) return null
  return releaseDetached(rec)
}

/** true while a torn-off window is following the shell's held pointer */
export function isTearingOff(): boolean {
  return torn !== null && !torn.rec.window.isDestroyed()
}
