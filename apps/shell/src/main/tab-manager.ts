import { basename, join } from 'node:path'
import { realpathSync } from 'node:fs'
import { app, BrowserWindow, shell, WebContentsView } from 'electron'
import type { Rectangle, WebContents } from 'electron'

import { appForKind, appIsDirty, documentApps, type AppCloseGuard } from './app-registry'
import type { AiDocContent } from '../../../docs/src/shared/ipc'
import { rendererUrl, safeExternalUrl } from '@genoffice/electron-utils'
import { HOME_CHANNELS } from '../shared/home-api'
import type { DocumentTabKind, OpenDocumentTab, TabKind, TabSummary } from '../shared/tabs-api'
import { isSettledRect, tabStripHeight } from '../shared/tab-drag-geometry'
import {
  DEFAULT_IDLE_MS,
  defaultMaxLiveTabs,
  MEMORY_BUDGET_SHARE,
  pickTabsToSleep,
  SLEEP_SWEEP_INTERVAL_MS,
  type SleepPolicy,
} from './tab-hibernation'

/** a tab lifted out of the strip with its live view: what "Open in New Window",
 *  tear-off and dock hand back and forth between the shell and a detached window */
export interface DetachedTab {
  view: WebContentsView
  kind: TabKind
  title: string
  filePath?: string
}

interface TabRecord {
  id: string
  kind: TabKind
  /** null for the Home tab — it's rendered by the shell window's own webContents */
  view: WebContentsView | null
  title: string
  filePath?: string
  /** chrome-free Present tab: no file, no editor menu or save/export targets */
  present?: boolean
  /** when the tab last became or stopped being the active one */
  lastActiveAt: number
  /** never put to sleep (tab context menu) */
  pinned?: boolean
  /** renderer destroyed to free memory; `view` is null until the tab is shown again */
  dormant?: boolean
}

const HOME_ID = 'home'

/**
 * The palette's Ctrl/Cmd+K summon, shared by the shell's tab views and the
 * detached editor windows so the same app behaves the same in both. The chord
 * is only claimed plain: shift/alt variants always belong to the app beneath
 * (Word-style dialog shortcuts), and `ownsPlainCtrlK` lets an app that ships
 * its own plain binding (docs: Word's hyperlink dialog) keep it — the palette
 * stays reachable through the global hotkey and the tray either way. Control
 * combos arrive as rawKeyDown on Windows, keyDown elsewhere. Returns the
 * listener's disposer.
 */
export function installPaletteSummon(
  wc: WebContents,
  ownsPlainCtrlK: () => boolean,
  summon: () => void,
): () => void {
  const listener = (event: Electron.Event, input: Electron.Input): void => {
    if (input.type !== 'keyDown' && input.type !== 'rawKeyDown') return
    if (!input.control && !input.meta) return
    if (input.shift || input.alt) return
    if (input.key?.toLowerCase() !== 'k') return
    if (ownsPlainCtrlK()) return
    event.preventDefault()
    summon()
  }
  wc.on('before-input-event', listener)
  return () => {
    if (!wc.isDestroyed()) wc.removeListener('before-input-event', listener)
  }
}

/** the live shell tab manager. main/index.ts owns its instance but hands it to
 *  no one; feature IPC modules registered without injected deps (the template
 *  center) reach the tab strip through this instead of a closure there. */
let liveTabManager: { manager: TabManager; window: BrowserWindow } | null = null

/** the shell's current tab manager, or null while the shell window is gone */
export function shellTabManager(): TabManager | null {
  const live = liveTabManager
  return live && !live.window.isDestroyed() ? live.manager : null
}

/**
 * Owns every open tab (Home + docs + sheets) inside the shell's single
 * BrowserWindow. Docs/sheets tabs are WebContentsView children of that
 * window; only the active one is visible at a time. Home has no view of its
 * own — hiding every other tab reveals the shell window's own content.
 */
export class TabManager {
  private readonly tabs: TabRecord[] = [
    { id: HOME_ID, kind: 'home', view: null, title: 'GenOffice', lastActiveAt: Date.now() },
  ]
  private activeId: string = HOME_ID
  private nextId = 1
  /** tab whose page entered HTML fullscreen (e.g. slides slideshow) — its view covers the tab strip */
  private htmlFullScreenId: string | null = null
  /** webContents ids whose view must cover the tab strip without HTML fullscreen
   *  (slides show: the window snaps via simpleFullScreen and asks for the bleed
   *  over IPC, since requestFullscreen would animate the native transition) */
  private readonly bleedWcIds = new Set<number>()
  /** tabs mid unsaved-changes prompt, so a second close click doesn't stack dialogs */
  private readonly closingIds = new Set<string>()
  private closedListener: ((tab: { kind: TabKind; filePath: string }) => void) | null = null
  /** views whose HTML-fullscreen listeners are installed: a view that leaves
   *  for a detached window and docks back must not get a second pair */
  private readonly fullScreenTracked = new WeakSet<WebContentsView>()
  /** pre-warmed view of the app that opens slowest (the sheets renderer: parsing
   *  its bundle and booting Univer is the bulk of a workbook's open time, and the
   *  shell hands the path over after mount anyway) */
  private spareView: WebContentsView | null = null
  private spareViewTimer: ReturnType<typeof setTimeout> | null = null
  /** views that already summon the command palette on Ctrl/Cmd+K */
  private readonly paletteHooked = new WeakSet<WebContents>()

  /**
   * The command palette's global summon: Ctrl/Cmd+K in any of the shell
   * window's surfaces (the home chrome or a document view) lands you on Home
   * with the palette open. Attached per webContents — `web-contents-created`
   * does not fire for this app's tab views, so activation is the hook point.
   * The lookup is live: a view torn into a detached window has no tab record
   * here and this listener stands down (the detached window's own hook is the
   * one alive); docked back, the record exists again and this one answers.
   */
  private installPaletteShortcut(wc: WebContents): void {
    if (this.paletteHooked.has(wc) || wc.isDestroyed()) return
    this.paletteHooked.add(wc)
    // the shell window's own webContents is Home's chrome: the palette host,
    // never an app that could own the chord
    const isShellChrome = wc === this.shellWindow.webContents
    const tabOf = (): TabRecord | undefined => this.tabs.find((t) => t.view?.webContents === wc)
    installPaletteSummon(
      wc,
      () => {
        if (isShellChrome) return false
        const tab = tabOf()
        return tab ? appForKind(tab.kind)?.ownsPlainCtrlK === true : false
      },
      () => {
        // a detached view is the detached window's to summon
        if (isShellChrome || tabOf()) this.summonPalette()
      },
    )
  }

  /** bring the shell window forward on Home and open the command palette —
   *  the shared Ctrl/Cmd+K answer, also used by the detached windows */
  summonPalette(): void {
    const win = this.shellWindow
    if (win.isDestroyed()) return
    if (win.isMinimized()) win.restore()
    win.show()
    win.focus()
    this.openHomeTab()
    win.webContents.send(HOME_CHANNELS.paletteOpen)
  }

  constructor(
    private readonly shellWindow: BrowserWindow,
    private readonly onChanged: () => void,
    private readonly applyMenuFor: (kind: TabKind) => void,
    /** localized placeholder title for a tab that has no file yet */
    private readonly untitledTitleFor?: (kind: TabKind) => string,
    /**
     * The live interface scale (genoffice#1913). A getter rather than a value so a change
     * made from the View menu reaches the next layout without rebuilding this.
     */
    private readonly uiScale: () => number = () => 1,
  ) {
    // Layout once synchronously for macOS/Windows (bounds are already correct),
    // then once more on the next tick. On Linux/X11, `resize` fires before the
    // window manager applies the new size, so getContentBounds() is still the
    // pre-maximize size inside the handler and a follow-up layout is required.
    // See https://github.com/genspark-ai/genoffice/issues/15
    shellWindow.on('resize', () => {
      this.layout()
      setImmediate(() => this.layout())
    })
    // Windows fires `restore` when a minimized window comes back from the taskbar,
    // sometimes before the client bounds have settled — the same stale-geometry
    // race as Linux/X11 maximize (issue #15). Reconcile synchronously (the bounds
    // are usually already right) and once more on the next tick. Bounded: exactly
    // two layout passes per restore, no polling, no recurring timers.
    shellWindow.on('restore', () => {
      this.reconcileAfterWindowRestore()
      setImmediate(() => this.reconcileAfterWindowRestore())
    })
    liveTabManager = { manager: this, window: shellWindow }
    this.installPaletteShortcut(shellWindow.webContents)
  }

  private scheduleSpareView(delayMs: number): void {
    if (process.env.GENOFFICE_NO_SPARE_VIEW || this.spareViewTimer || this.spareView) return
    if (!appForKind(this.tabs.find((t) => t.id === this.activeId)?.kind ?? 'home')?.open?.spare)
      return
    this.spareViewTimer = setTimeout(() => {
      this.spareViewTimer = null
      if (this.spareView || this.shellWindow.isDestroyed()) return
      const active = this.tabs.find((t) => t.id === this.activeId)
      const descriptor = active ? appForKind(active.kind) : undefined
      if (!active || !descriptor?.open?.spare) return
      const view = descriptor.createView()
      // registering the session made the spare the menu-action target
      descriptor.claimActiveTarget?.(() => active.view?.webContents ?? null)
      this.shellWindow.contentView.addChildView(view)
      view.setVisible(false)
      // the window may be minimized or mid-restore when the spare spawns; a
      // degenerate rect would leave the spare zero-sized until its next layout
      const bounds = this.contentBounds()
      if (isSettledRect(bounds)) view.setBounds(bounds)
      view.webContents.once('render-process-gone', () => {
        if (this.spareView !== view) return
        this.spareView = null
        view.webContents.close()
      })
      this.spareView = view
    }, delayMs)
  }

  private takeSpareView(): WebContentsView | null {
    const view = this.spareView
    this.spareView = null
    return view && !view.webContents.isDestroyed() ? view : null
  }

  private discardSpareView(): void {
    if (this.spareViewTimer) clearTimeout(this.spareViewTimer)
    this.spareViewTimer = null
    const spare = this.spareView
    this.spareView = null
    if (!spare) return
    this.shellWindow.contentView.removeChildView(spare)
    spare.webContents.close()
  }

  private untitled(kind: TabKind, fallback: string): string {
    return this.untitledTitleFor?.(kind) ?? fallback
  }

  private contentBounds(): Rectangle {
    const { width, height } = this.shellWindow.getContentBounds()
    if (this.htmlFullScreenId !== null && this.htmlFullScreenId === this.activeId) {
      return { x: 0, y: 0, width, height }
    }
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (active?.view && this.bleedWcIds.has(active.view.webContents.id)) {
      return { x: 0, y: 0, width, height }
    }
    // The strip is `zoom`ed by the interface scale, so its rendered height
    // grows with it — positioning the view at the unscaled constant would paint
    // the document underneath a taller tab strip.
    const strip = tabStripHeight(this.uiScale())
    return { x: 0, y: strip, width, height: Math.max(0, height - strip) }
  }

  /** Grow/restore a tab view over the tab strip on request (slides show fullscreen) */
  setContentBleed(wc: WebContents, on: boolean): void {
    if (on) this.bleedWcIds.add(wc.id)
    else this.bleedWcIds.delete(wc.id)
    this.layout()
  }

  /**
   * When a tab's page enters HTML fullscreen (the slides slideshow calls requestFullscreen),
   * grow its view over the tab strip so nothing of the shell chrome shows;
   * restore the normal bounds on leave.
   */
  private trackHtmlFullScreen(view: WebContentsView): void {
    if (this.fullScreenTracked.has(view)) return
    this.fullScreenTracked.add(view)
    // resolved at event time: the same view gets a fresh id every time it docks
    const currentId = () => this.tabs.find((t) => t.view === view)?.id
    view.webContents.on('enter-html-full-screen', () => {
      const id = currentId()
      if (id === undefined) return
      this.htmlFullScreenId = id
      this.layout()
    })
    view.webContents.on('leave-html-full-screen', () => {
      const id = currentId()
      if (id !== undefined && this.htmlFullScreenId === id) this.htmlFullScreenId = null
      this.layout()
    })
  }

  /** re-fit the active tab's view after a window resize */
  layout(): void {
    // Deferred resize layouts can land after the shell window was closed.
    if (this.shellWindow.isDestroyed()) return
    const bounds = this.contentBounds()
    // A minimized or mid-restore window can report a degenerate client rect;
    // applying it would shrink the active view to nothing — the white screen
    // after a Windows taskbar restore. Keep the last good layout instead.
    if (!isSettledRect(bounds)) return
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (active?.view) active.view.setBounds(bounds)
  }

  /**
   * Re-fit the active view after the native window returns from the taskbar.
   * Deliberately narrower than activateTab: no visibility toggles, no focus,
   * no menu, no onChanged — the active tab, its WebContentsView and every
   * other document stay exactly as they were.
   */
  reconcileAfterWindowRestore(): void {
    if (this.shellWindow.isDestroyed() || this.shellWindow.isMinimized()) return
    this.layout()
  }

  /** files open in any tab, for the open-documents registry */
  openFilePaths(): string[] {
    return this.tabs.flatMap((t) => (t.filePath ? [t.filePath] : []))
  }

  list(): TabSummary[] {
    return this.tabs.map((t) => ({
      id: t.id,
      kind: t.kind,
      title: t.title,
      closable: t.id !== HOME_ID,
      active: t.id === this.activeId,
      ...(t.filePath ? { filePath: t.filePath } : {}),
      ...(t.dormant ? { dormant: true } : {}),
      ...(t.pinned ? { pinned: true } : {}),
    }))
  }

  /**
   * Documents an MCP agent may act on: every editor tab except Home (no file)
   * and chrome-free Present tabs (a live preview of another tab's document, so
   * acting on it would double-count that document).
   *
   * Dirtiness is resolved here per family because it is not uniform — five
   * families answer synchronously in the main process, docs has to ask its
   * renderer. Hence the async signature.
   */
  async openDocuments(): Promise<OpenDocumentTab[]> {
    const tabs = this.tabs.filter(
      (tab) => tab.kind !== 'home' && !tab.present && (tab.view || tab.dormant),
    )
    return Promise.all(
      tabs.map(async (tab) => ({
        id: tab.id,
        kind: tab.kind as DocumentTabKind,
        title: tab.title,
        ...(tab.filePath ? { filePath: tab.filePath } : {}),
        active: tab.id === this.activeId,
        dirty: await this.tabIsDirty(tab),
      })),
    )
  }

  /** unsaved-changes state of one tab, whichever family owns it (docs asks its renderer) */
  private async tabIsDirty(tab: TabRecord): Promise<boolean> {
    const wc = tab.view?.webContents
    if (!wc || wc.isDestroyed()) return false
    return appIsDirty(tab.kind, wc)
  }

  openHomeTab(): void {
    this.activateTab(HOME_ID)
  }

  /** The in-app user manual (issue #1520): a closable tab showing the shell
      renderer's help mode — the same bundle, branched on ?mode=help, so no
      second renderer build is needed. */
  openHelpTab(): string {
    // F1 again, or the Help menu while the manual is already open, focuses it
    // instead of stacking a second copy: the tab is closable (only Home is
    // not), so "press F1 again" is an ordinary thing to do
    const existing = this.tabs.find((t) => t.kind === 'help')
    if (existing) {
      this.activateTab(existing.id)
      return existing.id
    }
    const view = createHelpView()
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.tabs.push({
      id,
      kind: 'help',
      view,
      title: this.untitled('help', 'GenOffice Help'),
      lastActiveAt: Date.now(),
    })
    this.activateTab(id)
    return id
  }

  /**
   * Open a tab of any registered app. The descriptor supplies the view factory
   * and the app's own open steps; everything else — child view, fullscreen
   * tracking, the strip record, activation — is app-agnostic. The per-kind
   * `open*Tab` methods below are the app-flavored entry points callers use.
   */
  openDocumentTab(
    kind: DocumentTabKind,
    openPath?: string,
    options?: { newBlank?: boolean; aiContent?: AiDocContent },
  ): string {
    const descriptor = appForKind(kind)
    if (!descriptor) throw new Error(`no registered app for "${kind}" tabs`)
    if (options?.newBlank) descriptor.open?.prepare?.(options)
    const spare = descriptor.open?.spare ? this.takeSpareView() : null
    const view = spare ?? descriptor.createView(openPath)
    if (spare && openPath) descriptor.open?.spare?.adopt?.(view, openPath)
    if (options) descriptor.open?.afterCreate?.(view.webContents.id, options)
    const id = `t${this.nextId++}`
    if (!spare) {
      this.shellWindow.contentView.addChildView(view)
      view.setVisible(false)
    }
    this.trackHtmlFullScreen(view)
    this.tabs.push({
      id,
      lastActiveAt: Date.now(),
      kind,
      view,
      title: openPath ? basename(openPath) : this.untitled(kind, descriptor.untitledTitle),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openDocsTab(
    openPath?: string,
    options?: { newBlank?: boolean; aiContent?: AiDocContent },
  ): string {
    return this.openDocumentTab('docs', openPath, options)
  }

  openSheetsTab(openPath?: string, options?: { newBlank?: boolean }): string {
    return this.openDocumentTab('sheets', openPath, options)
  }

  openSlidesTab(openPath?: string): string {
    return this.openDocumentTab('slides', openPath)
  }

  openPdfTab(openPath: string): string {
    return this.openDocumentTab('pdf', openPath)
  }

  /** Remount the tab's renderer so it re-reads its file from disk (View > Reload). */
  reloadTab(id: string): void {
    const tab = this.tabs.find((t) => t.id === id)
    const wc = tab?.view?.webContents
    if (!wc || wc.isDestroyed()) return
    appForKind(tab.kind)?.clearDirty?.(wc)
    wc.reload()
  }

  openMarkdownTab(openPath?: string): string {
    return this.openDocumentTab('markdown', openPath)
  }

  openHtmlTab(openPath?: string): string {
    return this.openDocumentTab('html', openPath)
  }

  /** Present → New tab: a chrome-free html tab showing the owner tab's live preview */
  openHtmlPresentTab(owner: WebContents, title: string): string {
    const descriptor = appForKind('html')!
    const view = descriptor.createPresentView!(owner, title)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(view)
    this.tabs.push({
      id,
      lastActiveAt: Date.now(),
      kind: 'html',
      view,
      title: title || this.untitled('html', descriptor.untitledTitle),
      present: true,
    })
    this.activateTab(id)
    return id
  }

  activateTab(id: string): void {
    const target = this.tabs.find((t) => t.id === id)
    if (!target) return
    if (target.dormant) this.reviveTab(target)
    const now = Date.now()
    const previous = this.tabs.find((t) => t.id === this.activeId)
    if (previous) previous.lastActiveAt = now
    target.lastActiveAt = now
    for (const t of this.tabs) t.view?.setVisible(t.id === id)
    const bounds = this.contentBounds()
    // a bring-forward while the window is minimized or mid-restore must not
    // zero-size the view being activated (same guard as layout())
    if (target.view && isSettledRect(bounds)) target.view.setBounds(bounds)
    this.activeId = id
    if (target.view) this.installPaletteShortcut(target.view.webContents)
    if (appForKind(target.kind)?.open?.spare) this.scheduleSpareView(3000)
    else this.discardSpareView()
    this.refreshActiveTargets()
    this.focusActiveView()
    this.onChanged()
  }

  /** Hand keyboard focus to the active tab's view. The click that opened or
   *  switched a tab lands on the chrome webContents (tab strip, Home list),
   *  and a WebContentsView made visible does not take focus on its own — so
   *  without this, typing after every open/switch keeps going to a hidden
   *  view. Home has no view of its own; the shell window's webContents is
   *  the focus target there. Skipped while the window is unfocused
   *  (background opens must not steal OS focus); the window's `focus`
   *  handler re-runs it. */
  focusActiveView(): void {
    if (this.shellWindow.isDestroyed() || !this.shellWindow.isFocused()) return
    const target = this.tabs.find((t) => t.id === this.activeId)
    if (!target) return
    if (target.view) target.view.webContents.focus()
    else this.shellWindow.webContents.focus()
  }

  /** Re-point the process-global active-editor targets and the app menu at this
   *  window's active tab. Called on every activation and on shell-window focus:
   *  a detached editor window ("Open in New Window") claims the same globals
   *  while it is focused. */
  refreshActiveTargets(): void {
    const target = this.tabs.find((t) => t.id === this.activeId)
    if (!target) return
    for (const descriptor of documentApps) {
      descriptor.claimActiveTarget?.(
        descriptor.id === target.kind && target.view ? () => target.view!.webContents : () => null,
      )
    }
    this.applyMenuFor(target.present ? 'home' : target.kind)
  }

  /** move a tab to a new index in the strip; Home is pinned at index 0 */
  reorderTab(id: string, toIndex: number): void {
    if (id === HOME_ID) return
    const fromIndex = this.tabs.findIndex((t) => t.id === id)
    if (fromIndex < 0) return
    const clamped = Math.min(Math.max(Math.trunc(toIndex), 1), this.tabs.length - 1)
    if (clamped === fromIndex) return
    const [moved] = this.tabs.splice(fromIndex, 1)
    this.tabs.splice(clamped, 0, moved)
    this.onChanged()
  }

  /** kind, title and file of the document tab rendered by `webContentsId` (Home has no view) */
  describeWebContents(
    webContentsId: number,
  ): { kind: TabKind; title: string; filePath?: string } | null {
    const tab = this.tabs.find((t) => t.view?.webContents.id === webContentsId)
    if (!tab) return null
    return { kind: tab.kind, title: tab.title, ...(tab.filePath ? { filePath: tab.filePath } : {}) }
  }

  tabIdForWebContents(webContentsId: number): string | undefined {
    return this.tabs.find((t) => t.view?.webContents.id === webContentsId)?.id
  }

  /** a module opened a file inside an existing tab (⌘O / queued path) — sync title + dedupe path */
  setTabFileFor(webContentsId: number, filePath: string): void {
    const tab = this.tabs.find((t) => t.view?.webContents.id === webContentsId)
    if (!tab) return
    tab.filePath = filePath
    tab.title = basename(filePath)
    this.onChanged()
  }

  /** an untitled document named itself before its first save (html: from the first AI request) */
  setTabTitleFor(webContentsId: number, title: string): void {
    const tab = this.tabs.find((t) => t.view?.webContents.id === webContentsId)
    if (!tab || tab.filePath || tab.title === title) return
    tab.title = title
    this.onChanged()
  }

  /** a file was renamed on disk (rename from the Home list) — sync any open tab's title/path;
   *  returns the affected views so callers can notify the embedded editors */
  renameTabFile(
    oldPath: string,
    newPath: string,
  ): Array<{ kind: TabKind; webContents: WebContents }> {
    const affected: Array<{ kind: TabKind; webContents: WebContents }> = []
    let renamed = 0
    for (const tab of this.tabs) {
      if (tab.filePath !== oldPath) continue
      tab.filePath = newPath
      tab.title = basename(newPath)
      renamed++
      if (tab.view) affected.push({ kind: tab.kind, webContents: tab.view.webContents })
    }
    if (renamed > 0) this.onChanged()
    return affected
  }

  /** live tabs of one app kind (MCP bridges and the close guard resolve their targets through this) */
  tabsOfKind(kind: DocumentTabKind): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === kind && t.view)
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** tabs of one app kind whose module reports unsaved changes (shell-close guard) */
  dirtyTabsOfKind(kind: DocumentTabKind): Array<{ id: string; webContents: WebContents }> {
    const isDirty = appForKind(kind)?.isDirty
    if (!isDirty) return []
    return this.tabs
      .filter((t) => t.kind === kind && t.view && isDirty(t.view.webContents))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** sheets tabs whose renderer reports unsaved journal edits (shell-close guard) */
  dirtySheetsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.dirtyTabsOfKind('sheets')
  }

  /** pdf tabs whose renderer reports unsaved markups/form edits (shell-close guard) */
  dirtyPdfTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.dirtyTabsOfKind('pdf')
  }

  /** markdown tabs whose renderer reports unsaved edits (shell-close guard) */
  dirtyMarkdownTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.dirtyTabsOfKind('markdown')
  }

  /** html tabs whose renderer reports unsaved edits (shell-close guard) */
  dirtyHtmlTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.dirtyTabsOfKind('html')
  }

  /** slides tabs whose main-process session has unsaved edits (shell-close guard) */
  dirtySlidesTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.dirtyTabsOfKind('slides')
  }

  /** all live docs tabs — dirtiness lives renderer-side, caller queries async (shell-close guard) */
  docsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabsOfKind('docs')
  }

  /** all live slides tabs (MCP bridge resolves its new tab's webContents through this) */
  slidesTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabsOfKind('slides')
  }

  /** all live sheets tabs (MCP bridge resolves its new tab's webContents through this) */
  sheetsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabsOfKind('sheets')
  }

  /** closes whichever tab is currently active; no-op for Home (Cmd+W target) */
  closeActiveTab(): void {
    void this.closeTab(this.activeId)
  }

  async closeTab(id: string): Promise<void> {
    if (id === HOME_ID) return
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab || this.closingIds.has(id)) return
    const descriptor = appForKind(tab.kind)
    let closeGuard: AppCloseGuard | null = null
    if (descriptor && tab.view) {
      if (descriptor.queryDirty) {
        // docs dirty state lives in the renderer and needs an async query; skip
        // the guard when clean (avoids a flash activation)
        this.closingIds.add(id)
        try {
          if (await descriptor.queryDirty(tab.view.webContents))
            closeGuard = descriptor.requestClose
        } finally {
          this.closingIds.delete(id)
        }
      } else if (descriptor.isDirty?.(tab.view.webContents)) {
        closeGuard = descriptor.requestClose
      }
    }
    if (closeGuard && tab.view) {
      // Bring the tab into view so the save prompt has visible context.
      if (this.activeId !== id) this.activateTab(id)
      this.closingIds.add(id)
      try {
        if (!(await closeGuard(tab.view.webContents, this.shellWindow))) return
      } finally {
        this.closingIds.delete(id)
      }
    }
    this.removeTab(id)
  }

  /**
   * Close a tab with no save prompt. The caller must have already settled the
   * document's unsaved changes (the MCP close tool saves or discards first):
   * this is the plain removal step of `closeTab`, so a dialog never appears in
   * a flow the user did not start.
   * Returns false when there is no such tab (already closed) or one is mid-prompt.
   */
  closeTabWithoutPrompt(id: string): boolean {
    if (id === HOME_ID) return false
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab || this.closingIds.has(id)) return false
    this.removeTab(id)
    return true
  }

  /** detach + drop one tab and re-activate a neighbour (no guards, no prompts) */
  private removeTab(id: string): void {
    const idx = this.tabs.findIndex((t) => t.id === id)
    if (idx < 0) return
    if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
    const [removed] = this.tabs.splice(idx, 1)
    if (removed.filePath && !removed.present)
      this.closedListener?.({ kind: removed.kind, filePath: removed.filePath })
    if (this.activeId === id) {
      const fallback = this.tabs[idx - 1] ?? this.tabs[0]
      this.activateTab(fallback.id)
    } else {
      this.onChanged()
    }
    if (removed.view) {
      removed.view.setVisible(false)
      this.shellWindow.contentView.removeChildView(removed.view)
      destroyView(removed.kind, removed.view)
    }
  }

  /** a tab with a renderer; dormant and Home tabs have none */
  private isLive(tab: TabRecord): boolean {
    return !!tab.view
  }

  setPinned(id: string, pinned: boolean): void {
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab || tab.kind === 'home') return
    tab.pinned = pinned
    this.onChanged()
  }

  /**
   * Destroy the renderer of a background tab and keep only what is needed to
   * reopen it: kind, path, title, position. Refuses the active tab, pinned
   * tabs, untitled documents, Present tabs, the manual, a tab mid-prompt or in
   * HTML fullscreen, and anything with unsaved changes. Resolves true when the
   * tab went to sleep.
   */
  async sleepTab(id: string): Promise<boolean> {
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab?.view || !this.canSleep(tab)) return false
    if (await this.tabIsDirty(tab)) return false
    // re-check: the dirty query yielded, and the user may have switched here meanwhile
    if (!tab.view || !this.canSleep(tab)) return false
    const view = tab.view
    tab.view = null
    tab.dormant = true
    view.setVisible(false)
    this.shellWindow.contentView.removeChildView(view)
    destroyView(tab.kind, view)
    this.onChanged()
    return true
  }

  private canSleep(tab: TabRecord): boolean {
    return (
      tab.id !== this.activeId &&
      !tab.pinned &&
      !tab.present &&
      !!tab.filePath &&
      tab.kind !== 'home' &&
      tab.kind !== 'help' &&
      !this.closingIds.has(tab.id) &&
      this.htmlFullScreenId !== tab.id
    )
  }

  /** rebuild a sleeping tab's view from its file, in place; the caller shows it */
  private reviveTab(tab: TabRecord): void {
    const view = tab.filePath ? appForKind(tab.kind)?.createView(tab.filePath) : undefined
    if (!view) return
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(view)
    tab.view = view
    tab.dormant = false
    tab.lastActiveAt = Date.now()
  }

  /** a sleeping tab asked for by its webContents (agents, CLI, menus) wakes in the background */
  private ensureLive(tab: TabRecord | undefined): TabRecord | undefined {
    if (tab?.dormant) {
      this.reviveTab(tab)
      this.onChanged()
    }
    return tab
  }

  /**
   * One pass of the sleep policy: background tabs idle past `policy.idleMs`
   * sleep while more than `policy.maxLiveTabs` renderers are alive; over the
   * memory budget the oldest sleeps regardless. Returns the ids put to sleep.
   */
  async sleepIdleTabs(
    policy: SleepPolicy,
    now = Date.now(),
    overBudget = false,
  ): Promise<string[]> {
    const live = this.tabs.filter((t) => this.isLive(t))
    const candidates: { id: string; lastActiveAt: number }[] = []
    // dirtiness is resolved here, not in canSleep: it is async for docs, and a
    // dirty tab left among the candidates would be picked (and refused) on
    // every sweep, starving the clean tabs behind it
    for (const t of live) {
      if (this.canSleep(t) && !(await this.tabIsDirty(t)))
        candidates.push({ id: t.id, lastActiveAt: t.lastActiveAt })
    }
    const slept: string[] = []
    for (const id of pickTabsToSleep(candidates, live.length, policy, now, overBudget)) {
      if (await this.sleepTab(id)) slept.push(id)
    }
    return slept
  }

  /**
   * Run the sleep policy every minute while `enabled()` holds. The live-tab
   * cap follows the machine's memory; the budget check sums every process of
   * the app against a share of physical memory.
   */
  startSleepSweeper(enabled: () => boolean): () => void {
    const policy: SleepPolicy = {
      maxLiveTabs:
        envNumber('GENOFFICE_TAB_SLEEP_MAX_LIVE') ??
        defaultMaxLiveTabs(process.getSystemMemoryInfo().total * 1024),
      idleMs: envNumber('GENOFFICE_TAB_SLEEP_IDLE_MS') ?? DEFAULT_IDLE_MS,
    }
    let running = false
    const timer = setInterval(
      () => {
        if (!enabled() || running || this.shellWindow.isDestroyed()) return
        running = true
        void this.sleepIdleTabs(policy, Date.now(), appOverMemoryBudget()).finally(() => {
          running = false
        })
      },
      envNumber('GENOFFICE_TAB_SLEEP_SWEEP_MS') ?? SLEEP_SWEEP_INTERVAL_MS,
    )
    timer.unref?.()
    return () => clearInterval(timer)
  }

  /** told when a tab with a file is closed (not detached), so it can be reopened */
  setClosedListener(listener: ((tab: { kind: TabKind; filePath: string }) => void) | null): void {
    this.closedListener = listener
  }

  /** Remove a tab from the strip WITHOUT destroying its WebContentsView and
   *  hand it to the caller ("Open in New Window" — the live document, unsaved
   *  edits included, moves into a detached editor window). Null while a close
   *  prompt is pending on the tab. */
  detachTab(id: string): DetachedTab | null {
    if (id === HOME_ID) return null
    const idx = this.tabs.findIndex((t) => t.id === id)
    const tab = idx >= 0 ? this.tabs[idx] : undefined
    if (!tab?.view || tab.present || this.closingIds.has(id)) return null
    this.tabs.splice(idx, 1)
    if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
    const view = tab.view
    view.setVisible(false)
    this.shellWindow.contentView.removeChildView(view)
    if (this.activeId === id) {
      const fallback = this.tabs[idx - 1] ?? this.tabs[0]
      this.activateTab(fallback.id)
    } else {
      this.onChanged()
    }
    return { view, kind: tab.kind, title: tab.title, filePath: tab.filePath }
  }

  /** Whether a tab can leave the strip for its own window (tear-off / Open in
   *  New Window): every document tab except a chrome-free Present tab. */
  canDetachTab(id: string): boolean {
    const tab = this.tabs.find((t) => t.id === id)
    return !!tab?.view && !tab.present && !this.closingIds.has(id)
  }

  /**
   * The inverse of detachTab: a live view coming back from a detached window
   * (dock by dragging the window onto the strip, or a tear-off that returned
   * mid-gesture) becomes a tab again at `index` — Home stays pinned at 0, an
   * out-of-range or omitted index appends — and is activated. The document,
   * unsaved edits included, is untouched: only the view's parent changes.
   */
  attachTab(record: DetachedTab, index?: number): string {
    const id = `t${this.nextId++}`
    const { view } = record
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(view)
    const slot =
      index === undefined || !Number.isFinite(index)
        ? this.tabs.length
        : Math.min(Math.max(Math.trunc(index), 1), this.tabs.length)
    this.tabs.splice(slot, 0, {
      id,
      kind: record.kind,
      view,
      title: record.title,
      filePath: record.filePath,
      lastActiveAt: Date.now(),
    })
    this.activateTab(id)
    return id
  }

  /** the editor tab showing this file, whichever app owns it (path compared after resolving links) */
  findTabByPath(
    path?: string,
  ): { id: string; kind: TabKind; webContents: WebContents } | undefined {
    const wanted = canonicalPath(path)
    const tab = this.ensureLive(
      this.tabs.find(
        (t) =>
          (t.view || t.dormant) && t.filePath && !t.present && canonicalPath(t.filePath) === wanted,
      ),
    )
    return tab?.view ? { id: tab.id, kind: tab.kind, webContents: tab.view.webContents } : undefined
  }

  webContentsForTab(id: string): WebContents | undefined {
    return this.ensureLive(this.tabs.find((t) => t.id === id))?.view?.webContents
  }

  /** the tab of one app kind showing this file (the open route's dedupe, per kind) */
  findTabOfKindByPath(kind: DocumentTabKind, path?: string): string | undefined {
    const wanted = canonicalPath(path)
    return this.tabs.find(
      (t) =>
        t.kind === kind &&
        (t.view || t.dormant) &&
        t.filePath &&
        !t.present &&
        canonicalPath(t.filePath) === wanted,
    )?.id
  }

  findDocsTabByPath(path?: string): string | undefined {
    return this.findTabOfKindByPath('docs', path)
  }

  findSheetsTab(): string | undefined {
    return this.tabs.find((t) => t.kind === 'sheets')?.id
  }

  findSheetsTabByPath(path?: string): string | undefined {
    return this.findTabOfKindByPath('sheets', path)
  }

  findSlidesTabByPath(path?: string): string | undefined {
    return this.findTabOfKindByPath('slides', path)
  }

  findPdfTabByPath(path?: string): string | undefined {
    return this.findTabOfKindByPath('pdf', path)
  }

  findMarkdownTabByPath(path?: string): string | undefined {
    return this.findTabOfKindByPath('markdown', path)
  }

  findHtmlTabByPath(path?: string): string | undefined {
    return this.findTabOfKindByPath('html', path)
  }

  /** the active tab's html view, if the active tab is html (html menu target) */
  activeHtmlTab(): { id: string; webContents: WebContents; filePath?: string } | undefined {
    const tab = this.tabs.find((t) => t.id === this.activeId)
    return tab?.kind === 'html' && tab.view && !tab.present
      ? { id: tab.id, webContents: tab.view.webContents, filePath: tab.filePath }
      : undefined
  }

  /** the active tab's markdown view, if the active tab is markdown (markdown menu target) */
  activeMarkdownTab(): { id: string; webContents: WebContents; filePath?: string } | undefined {
    const tab = this.tabs.find((t) => t.id === this.activeId)
    return tab?.kind === 'markdown' && tab.view
      ? { id: tab.id, webContents: tab.view.webContents, filePath: tab.filePath }
      : undefined
  }

  /** the active tab's pdf view, if the active tab is a pdf (pdf menu target) */
  activePdfTab(): { id: string; webContents: WebContents; filePath?: string } | undefined {
    const tab = this.tabs.find((t) => t.id === this.activeId)
    return tab?.kind === 'pdf' && tab.view
      ? { id: tab.id, webContents: tab.view.webContents, filePath: tab.filePath }
      : undefined
  }
}

/** release a view's renderer; an app whose teardown needs main-process bookkeeping
 *  (docs) gets it dropped before the close */
function destroyView(kind: TabKind, view: WebContentsView): void {
  if (view.webContents.isDestroyed()) return
  appForKind(kind)?.teardown?.(view.webContents)
  view.webContents.close()
}

/** test hooks: the sweep is minutes-scale by design, too slow for an e2e run */
function envNumber(name: string): number | undefined {
  const value = Number(process.env[name])
  return Number.isFinite(value) && value > 0 ? value : undefined
}

function appOverMemoryBudget(): boolean {
  try {
    const usedKB = app.getAppMetrics().reduce((sum, m) => sum + m.memory.workingSetSize, 0)
    return usedKB > process.getSystemMemoryInfo().total * MEMORY_BUDGET_SHARE
  } catch {
    return false
  }
}

export function canonicalPath(path: string | undefined): string | undefined {
  if (path === undefined) return undefined
  try {
    return realpathSync.native(path)
  } catch {
    return path
  }
}

/** The manual shares the shell's renderer bundle and preload: the view is just
    the shell renderer at ?mode=help. Dev serves it from the dev server, the
    packaged app from the `help` scheme root registered in shell main. */
function createHelpView(): WebContentsView {
  const view = new WebContentsView({
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })
  view.webContents.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url)
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  void view.webContents.loadURL(
    rendererUrl(process.env.ELECTRON_RENDERER_URL, 'help', { mode: 'help' }),
  )
  return view
}
