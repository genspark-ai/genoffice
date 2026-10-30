import { basename } from 'node:path'
import { BrowserWindow, ipcMain, screen, WebContentsView } from 'electron'
import type { Rectangle, WebContents } from 'electron'

import { appForKind, appIsDirty, resetSheetsShuttingDown } from './app-registry'
import { canonicalPath } from './tab-manager'
import type { DetachedTab } from './tab-manager'
import { shellQuitTracker } from './window-close-guard'
import {
  HOME_CHANNELS,
  type DesktopPinInfo,
  type DesktopPinKind,
  type DesktopPinRecord,
} from '../shared/home-api'
import type { OpenDocumentTab, TabKind } from '../shared/tabs-api'
import {
  DOCK_BAND_SLACK,
  DOCK_DWELL_MS,
  dockBand,
  isSettledRect,
  pointInRect,
} from '../shared/tab-drag-geometry'

/**
 * Detached editor windows: a tab's live WebContentsView reparented into its
 * own BrowserWindow, so the document — including unsaved edits — moves
 * without a reload. Three ways in, two ways out:
 *
 *   in:  "Open in New Window" on a tab; dragging a tab off the strip
 *        (tear-off: the new window follows the still-held pointer);
 *        "Pin to Desktop" — the same move, but the window becomes a
 *        borderless widget on the Home desktop (see the pinned-widget
 *        section below);
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
  /** a pinned desktop widget: frameless, strip-carrying, close = dock back */
  pinned: boolean
  /** the widget's drag strip (pinned windows only; owns the dock-back control) */
  strip?: WebContentsView
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

// ── pinned desktop widgets ─────────────────────────────
//
// "Pin to Desktop" turns a detached editor window into a live widget on the
// Home desktop: frameless, carrying a slim drag strip (the editor view has no
// drag region of its own), parented to the shell window so it always stacks
// above the desktop it belongs to, and shown without focus so pinning never
// disturbs what the person was typing. Closing the widget is a dock-back —
// the live view re-homes into the tab strip, edits included, so no prompt
// stands between the person and the document (a quit falls back to the
// unsaved-changes guard instead, exactly like any other detached window).
//
// Widgets are visible only while the desktop they decorate is: the renderer
// signals the desktop view's mount state, the shell signals whether the Home
// tab is the active one, and a widget shows only while both hold.

/** how the pin store is reached; injected by index.ts (desktop-pins.json) so
 *  this module stays Electron-free below the window layer and testable */
export interface DesktopPinStore {
  read(): DesktopPinRecord[]
  write(pins: readonly DesktopPinRecord[]): void
}

let pinStore: DesktopPinStore | null = null
let pinsChangedListener: (pins: DesktopPinInfo[]) => void = () => {}
/** the desktop view is mounted in the shell renderer */
let pinsRendererVisible = false
/** the Home tab is the active tab (another tab covers the desktop otherwise) */
let pinsHostActive = true
/** write-through window: a burst of widget moves costs one save */
const PIN_PERSIST_MS = 500
let pinPersistTimer: ReturnType<typeof setTimeout> | null = null

/** the editor kinds that can become a widget (mirrors DesktopPinKind) */
const PINNABLE_KINDS: ReadonlySet<string> = new Set([
  'docs',
  'sheets',
  'slides',
  'pdf',
  'markdown',
  'html',
])

/** wire the persistence backend; null detaches (tests, early shutdown) */
export function setDesktopPinStore(store: DesktopPinStore | null): void {
  pinStore = store
}

/** fires whenever a widget is pinned or unpinned (open-documents publishing style) */
export function setPinsChangedListener(listener: (pins: DesktopPinInfo[]) => void): void {
  pinsChangedListener = listener
}

/** the live widgets in the shape the renderer and the MCP host consume */
export function pinnedDesktopPins(): DesktopPinInfo[] {
  const out: DesktopPinInfo[] = []
  for (const [wcId, rec] of detached) {
    if (!rec.pinned || rec.window.isDestroyed() || !PINNABLE_KINDS.has(rec.kind)) continue
    const bounds = rec.window.getContentBounds()
    out.push({
      id: ID_PREFIX + wcId,
      kind: rec.kind as DesktopPinInfo['kind'],
      title: rec.window.getTitle(),
      filePath: rec.filePath ?? null,
      bounds: { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height },
    })
  }
  return out
}

/** the persisted snapshot: what to reopen and where; untitled documents cannot
 *  survive a restart, so they are absent here by design */
function pinSnapshot(): DesktopPinRecord[] {
  const out: DesktopPinRecord[] = []
  for (const rec of detached.values()) {
    if (!rec.pinned || rec.window.isDestroyed() || !rec.filePath) continue
    if (!PINNABLE_KINDS.has(rec.kind)) continue
    const bounds = rec.window.getContentBounds()
    out.push({
      kind: rec.kind as DesktopPinKind,
      filePath: rec.filePath,
      bounds: { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height },
    })
  }
  return out
}

function persistPinsNow(): void {
  if (pinPersistTimer) {
    clearTimeout(pinPersistTimer)
    pinPersistTimer = null
  }
  pinStore?.write(pinSnapshot())
}

/** one save per burst of moves/resizes; pinning or unpinning saves at once */
function schedulePersistPins(): void {
  if (!pinStore || pinPersistTimer) return
  pinPersistTimer = setTimeout(() => {
    pinPersistTimer = null
    persistPinsNow()
  }, PIN_PERSIST_MS)
}

function notifyPinsChanged(): void {
  pinsChangedListener(pinnedDesktopPins())
}

/** show/hide every widget: the desktop view and the Home tab must both be up */
function applyPinVisibility(): void {
  const show = pinsRendererVisible && pinsHostActive
  for (const rec of detached.values()) {
    if (!rec.pinned || rec.window.isDestroyed()) continue
    if (show) {
      if (!rec.window.isVisible()) rec.window.showInactive()
    } else if (rec.window.isVisible()) {
      rec.window.hide()
    }
  }
}

/** the renderer's desktop view mount state (pinned widgets follow it) */
export function setDesktopPinsVisible(visible: boolean): void {
  pinsRendererVisible = visible
  applyPinVisibility()
}

/** the surface widgets decorate: the Home tab (a covering tab hides them) */
export function setPinsHostSurfaceActive(active: boolean): void {
  pinsHostActive = active
  applyPinVisibility()
}

/** re-parent the widgets to a (re)created shell window — a child window keeps
 *  stacking above the desktop it belongs to */
export function attachPinnedWidgetsToShell(shell: BrowserWindow): void {
  if (shell.isDestroyed()) return
  for (const rec of detached.values()) {
    if (!rec.pinned || rec.window.isDestroyed()) continue
    rec.window.setParentWindow(shell)
  }
  applyPinVisibility()
}

/** the shell window is closing: unparent the widgets first, or they are
 *  destroyed with it and their unsaved edits die without a prompt */
export function detachPinnedWidgetsFromShell(shell: BrowserWindow): void {
  for (const rec of detached.values()) {
    if (!rec.pinned || rec.window.isDestroyed()) continue
    if (rec.window.getParentWindow() === shell) rec.window.setParentWindow(null)
  }
}

/** a document already pinned: show its widget instead of a second copy */
export function focusPinnedWidget(path: string): boolean {
  const wanted = canonicalPath(path)
  for (const rec of detached.values()) {
    if (
      rec.pinned &&
      rec.filePath &&
      canonicalPath(rec.filePath) === wanted &&
      !rec.window.isDestroyed()
    ) {
      rec.window.showInactive()
      return true
    }
  }
  return false
}

/** the widget IPC: the renderer's read/activate/visibility surface */
export function registerDetachedPinsIpc(): void {
  ipcMain.handle(HOME_CHANNELS.desktopPins, (): DesktopPinInfo[] => pinnedDesktopPins())
  ipcMain.handle(HOME_CHANNELS.desktopPinsActivate, (_event, rawId: unknown): boolean => {
    if (typeof rawId !== 'string') return false
    const rec = recordById(rawId)
    if (rec?.pinned && !rec.window.isDestroyed()) rec.window.showInactive()
    return rec?.pinned === true
  })
  ipcMain.handle(HOME_CHANNELS.desktopPinsSetVisible, (_event, raw: unknown): boolean => {
    if (typeof raw !== 'boolean') return false
    setDesktopPinsVisible(raw)
    return true
  })
}

// ── the widget's drag strip ────────────────────────────

/** strip height in window pixels; the editor view starts below it */
const PIN_STRIP_H = 28
/** the link target the strip's dock control opens (intercepted, never loaded) */
const PIN_DOCK_URL = 'genoffice-pin:dock'
/** default widget size: a compact document pane, resizable like any window */
const PIN_DEFAULT_SIZE = { width: 480, height: 360 }
const PIN_MIN_W = 320
const PIN_MIN_H = 240

/**
 * The strip's page: one line of HTML/CSS whose whole bar is a native drag
 * region and whose single control docks the document back. Colors follow the
 * tab-strip band values in title-bar-overlay.ts (--tabstrip-bg / text
 * equivalents) — this page cannot load tokens.css, so the two pairs are
 * duplicated here and must stay in sync with that file.
 */
const PIN_STRIP_HTML = `<!doctype html>
<meta charset="utf-8">
<meta name="color-scheme" content="light dark">
<style>
  html, body { margin: 0; height: 100%; font: 500 12px system-ui, sans-serif;
    -webkit-user-select: none; user-select: none; overflow: hidden; }
  .bar { display: flex; align-items: center; gap: 8px; height: 100%;
    padding: 0 10px 0 12px; background: #ebebeb; color: #454746;
    -webkit-app-region: drag; cursor: default; }
  .bar .title { flex: 1; min-width: 0; overflow: hidden;
    text-overflow: ellipsis; white-space: nowrap; }
  .bar .dock { -webkit-app-region: no-drag; text-decoration: none; color: inherit;
    font-size: 14px; line-height: 1; padding: 3px 7px; border-radius: 5px; }
  .bar .dock:hover { background: color-mix(in srgb, currentColor 14%, transparent); }
  @media (prefers-color-scheme: dark) {
    .bar { background: #2a2a2a; color: #e4e4e4; }
  }
</style>
<div class="bar"><span class="title">{title}</span>
  <a class="dock" href="${PIN_DOCK_URL}" target="_blank" title="⇤">&#8676;</a></div>`

function escapeHtml(text: string): string {
  return text.replace(/[&<>"']/g, (ch) => `&#${ch.codePointAt(0)};`)
}

/** build the strip view; `onDock` runs when its control is pressed */
function createPinStrip(title: string, onDock: () => void): WebContentsView {
  const strip = new WebContentsView({ webPreferences: { sandbox: true } })
  void strip.webContents.loadURL(
    'data:text/html;charset=utf-8,' +
      encodeURIComponent(PIN_STRIP_HTML.replace('{title}', escapeHtml(title))),
  )
  strip.webContents.setWindowOpenHandler(({ url }) => {
    if (url === PIN_DOCK_URL) onDock()
    return { action: 'deny' }
  })
  return strip
}

/** frame origin that puts a `slot`-th widget into the shell content's
 *  bottom-right corner, cascading inward; kept inside the content rect */
function pinWidgetBounds(parent: BrowserWindow, slot: number): Rectangle {
  const content = parent.getContentBounds()
  const step = 32
  const margin = 24
  const x = Math.max(
    content.x,
    content.x + content.width - PIN_DEFAULT_SIZE.width - margin - step * (slot % 4),
  )
  const y = Math.max(
    content.y,
    content.y + content.height - PIN_DEFAULT_SIZE.height - margin - step * Math.floor(slot / 4),
  )
  return { x: toInt(x), y: toInt(y), ...PIN_DEFAULT_SIZE }
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
export function detachedFileForWebContents(wcId: number): string | undefined {
  return detached.get(wcId)?.filePath
}

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
  return appIsDirty(rec.kind, wc)
}

/** run the family's save/don't-save/cancel prompt when the document is dirty;
 *  true when the window may close (same helpers as the tab close path) */
async function confirmDetachedClose(rec: DetachedRecord): Promise<boolean> {
  if (!(await detachedIsDirty(rec))) return true
  return (await appForKind(rec.kind)?.requestClose(rec.view.webContents, rec.window)) ?? true
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

/** a widget's strip label follows the document's name (Save As / rename) */
function refreshStripTitle(rec: DetachedRecord, title: string): void {
  const strip = rec.strip
  if (!strip || strip.webContents.isDestroyed()) return
  const text = JSON.stringify(title)
  void strip.webContents
    .executeJavaScript(`document.querySelector('.title').textContent = ${text}`)
    .catch(() => {
      // a strip torn down mid-call has nothing left to relabel
    })
}

/** Save As / open-in-place landed on a new path — keep the window title and
 *  the dedup registry in sync (same contract as TabManager.setTabFileFor). */
export function detachedSetFileFor(webContentsId: number, filePath: string): void {
  const rec = detached.get(webContentsId)
  if (!rec) return
  rec.filePath = filePath
  if (!rec.window.isDestroyed()) rec.window.setTitle(basename(filePath))
  if (rec.pinned) {
    // the pin's persisted identity is this path: follow it and relabel the strip
    refreshStripTitle(rec, basename(filePath))
    persistPinsNow()
  }
  onChanged()
}

/** an untitled document named itself before its first save (html: from the first AI request) */
export function detachedSetTitleFor(webContentsId: number, title: string): void {
  const rec = detached.get(webContentsId)
  if (!rec || rec.filePath || rec.window.isDestroyed()) return
  rec.window.setTitle(title)
  if (rec.pinned) refreshStripTitle(rec, title)
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
      if (rec.pinned) {
        refreshStripTitle(rec, basename(newPath))
        persistPinsNow()
      }
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

/** how many widgets exist right now (the cascade slot for the next one) */
function pinnedCount(): number {
  let n = 0
  for (const rec of detached.values()) if (rec.pinned && !rec.window.isDestroyed()) n += 1
  return n
}

/** dock a pinned widget back into the shell strip: the live view re-homes with
 *  its edits, so there is nothing to confirm and nothing to discard */
function dockPinnedWidget(rec: DetachedRecord): void {
  if (rec.released || rec.window.isDestroyed()) return
  if (!dockHost) {
    // no shell to dock into (it closed or is closing): the ordinary close path
    // takes over — the guard prompts if the document is dirty
    rec.window.close()
    return
  }
  dockDetachedWindow(rec)
}

/**
 * Lift the document out of its window without closing the renderer: the view
 * is unparented and the window destroyed (no close prompt — nothing is being
 * discarded). The caller re-homes the view.
 */
function releaseDetached(rec: DetachedRecord): DetachedTab {
  const { window: win, view, kind, filePath, strip } = rec
  clearDockHover(rec)
  rec.released = true
  if (torn?.rec === rec) torn = null
  const title = win.isDestroyed() ? (filePath ? basename(filePath) : '') : win.getTitle()
  detached.delete(view.webContents.id)
  if (strip) {
    win.contentView.removeChildView(strip)
    if (!strip.webContents.isDestroyed()) strip.webContents.close()
    rec.strip = undefined
  }
  if (!win.isDestroyed()) {
    win.contentView.removeChildView(view)
    win.destroy()
  }
  if (rec.pinned) {
    persistPinsNow()
    notifyPinsChanged()
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

/** screen rect of the shell's dock band for a given shell window (null when hidden).
 *  `scale` is the live interface scale: the band has to cover the strip as it is
 *  actually rendered, or a scaled-up strip outgrows the band and re-docking
 *  misses. */
export function dockBandFor(shellWindow: BrowserWindow | null, scale = 1): Rectangle | null {
  if (!shellWindow || shellWindow.isDestroyed() || !shellWindow.isVisible()) return null
  if (shellWindow.isMinimized()) return null
  return dockBand(shellWindow.getContentBounds(), DOCK_BAND_SLACK, scale)
}

/** frame origin that puts the pointer (at) inside the new window's title bar,
 *  kept on the pointer's display */
function tearOffOrigin(at: { x: number; y: number }, size: { width: number; height: number }) {
  const area = screen.getDisplayNearestPoint(at).workArea
  const x = Math.min(Math.max(at.x - TEAR_GRAB.dx, area.x), area.x + area.width - size.width)
  const y = Math.min(Math.max(at.y - TEAR_GRAB.dy, area.y), area.y + area.height - size.height)
  return { x: toInt(x), y: toInt(y) }
}

// Electron rejects -0 as an int argument ("conversion failure"); Math.round
// yields -0 for values in (-0.5, 0), e.g. a window steered to the screen edge
function toInt(v: number): number {
  return Math.round(v) || 0
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
  /**
   * pin to desktop: the window becomes a borderless widget over the Home
   * desktop instead of a framed editor window — parented to the shell window,
   * shown inactive, carrying a drag strip, closed by docking back
   */
  pin?: {
    /** the shell window the widget floats above (keeps its stacking) */
    parent: BrowserWindow
    /** the widget's screen rect; omitted = cascade the bottom-right corner */
    bounds?: Rectangle
  }
}): BrowserWindow {
  const { view, kind, title, filePath, applyMenuFor, tearOffAt, pin } = options
  const size = pin ? PIN_DEFAULT_SIZE : (DEFAULT_SIZE[kind] ?? { width: 1360, height: 900 })
  const origin = tearOffAt ? tearOffOrigin(tearOffAt, size) : {}
  const win = pin
    ? new BrowserWindow({
        ...(pin.bounds ?? pinWidgetBounds(pin.parent, pinnedCount())),
        minWidth: PIN_MIN_W,
        minHeight: PIN_MIN_H,
        title,
        show: false,
        // frameless: the widget's own drag strip replaces the OS title bar
        frame: false,
        resizable: true,
        // a child window always stacks above the shell — the widget stays on
        // its desktop no matter which of the two has focus
        parent: pin.parent,
        // a widget, not a window the taskbar should advertise
        skipTaskbar: true,
      })
    : new BrowserWindow({
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
    if (rec.strip) {
      win.contentView.removeChildView(rec.strip)
      if (!rec.strip.webContents.isDestroyed()) rec.strip.webContents.close()
      rec.strip = undefined
    }
    detached.delete(wcId)
    const descriptor = appForKind(kind)
    if (descriptor?.teardown) descriptor.teardown(view.webContents)
    else view.webContents.close()
    win.destroy()
    if (rec.pinned) {
      persistPinsNow()
      notifyPinsChanged()
    }
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
    pinned: !!pin,
  }
  detached.set(wcId, rec)
  onChanged()
  if (pin) {
    rec.strip = createPinStrip(title, () => {
      // the strip's dock control: same path as a close (no prompt — the live
      // view simply re-homes into the tab strip)
      if (!win.isDestroyed() && !rec.released) dockPinnedWidget(rec)
    })
    persistPinsNow()
    notifyPinsChanged()
  }

  win.contentView.addChildView(view)
  if (rec.strip) win.contentView.addChildView(rec.strip)
  const layout = () => {
    if (win.isDestroyed()) return
    const bounds = win.getContentBounds()
    // A minimized or mid-restore window can report a degenerate client rect;
    // applying it would blank the document view (the white screen after a
    // Windows taskbar restore). Keep the last good layout instead.
    if (!isSettledRect(bounds)) return
    const top = rec.pinned ? PIN_STRIP_H : 0
    view.setBounds({ x: 0, y: top, width: bounds.width, height: bounds.height - top })
    rec.strip?.setBounds({ x: 0, y: 0, width: bounds.width, height: PIN_STRIP_H })
  }
  // same Linux/X11 quirk as the shell window: `resize` fires before the WM
  // applies the new size, so a follow-up layout on the next tick is required
  win.on('resize', () => {
    layout()
    setImmediate(() => layout())
  })
  // a taskbar restore needs the same sync + next-tick reconciliation: on
  // Windows the client bounds may not have settled when `restore` fires
  win.on('restore', () => {
    layout()
    setImmediate(() => layout())
  })
  // a moved/resized widget re-persists its rect, debounced like the layout save
  if (pin) {
    win.on('moved', () => schedulePersistPins())
    win.on('resized', () => schedulePersistPins())
  }
  layout()
  view.setVisible(true)
  if (tearOffAt) {
    const pos = win.getPosition()
    torn = { rec, grabDx: tearOffAt.x - pos[0], grabDy: tearOffAt.y - pos[1] }
    win.showInactive()
  } else if (pin) {
    // pinning never disturbs what the person was typing: the widget appears
    // inactive, and only when the desktop view is actually showing it
    applyPinVisibility()
  } else {
    view.webContents.focus()
  }

  // the title is owned here (detach-time tab title, Save As updates);
  // the renderer's static <title> must not overwrite it
  win.on('page-title-updated', (event) => event.preventDefault())

  // The focused window owns the process-global menu-command targets; the
  // shell's tab manager re-claims both on its own window's focus.
  win.on('focus', () => {
    appForKind(kind)?.claimActiveTarget?.(() =>
      view.webContents.isDestroyed() ? null : view.webContents,
    )
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
  const vetoQuit = (): void => {
    resetSheetsShuttingDown()
    shellQuitTracker().reset()
  }
  win.on('close', (event) => {
    if (closeConfirmed || rec.released) return
    // a pinned widget's close means "back to the dock", not "discard": the
    // live view re-homes into the tab strip with its edits, so no prompt
    // stands between the person and the document. A quit (or a shell that is
    // already gone) falls back to the guard below — the pin is persisted, but
    // unsaved edits must never die silently.
    if (rec.pinned && !shellQuitTracker().quitting() && dockHost) {
      event.preventDefault()
      dockPinnedWidget(rec)
      return
    }
    event.preventDefault()
    if (promptInFlight) return
    promptInFlight = true
    void (async () => {
      try {
        // a denied close vetoes any quit that was in flight: the sheets close
        // guard must prompt again on later closes instead of silently proceeding
        if (await confirmDetachedClose(rec)) tearDown()
        else vetoQuit()
      } catch (err) {
        console.error('[shell] detached window close prompt failed:', err)
        vetoQuit()
      } finally {
        promptInFlight = false
      }
    })()
  })
  win.on('closed', () => {
    clearDockHover(rec)
    if (torn?.rec === rec) torn = null
    detached.delete(wcId)
    if (rec.pinned) {
      persistPinsNow()
      notifyPinsChanged()
    }
    onChanged()
  })
  return win
}

/** steer the torn-off window after the held pointer (screen coordinates) */
export function dragTornWindow(x: number, y: number): void {
  if (!torn || torn.rec.window.isDestroyed()) return
  torn.rec.window.setPosition(toInt(x - torn.grabDx), toInt(y - torn.grabDy))
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
