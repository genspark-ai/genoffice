import { basename } from 'node:path'
import { realpathSync } from 'node:fs'
import { BrowserWindow } from 'electron'
import type { Rectangle, WebContents, WebContentsView } from 'electron'

import {
  createDocsView,
  docsQueryDirty,
  markDocsNewBlank,
  queueDocsAiContent,
  requestDocsClose,
  setActiveDocsResolver,
  teardownDocsRenderer,
} from '../../../docs/src/main/docs-main'
import type { AiDocContent } from '../../../docs/src/shared/ipc'
import {
  createMarkdownView,
  markdownIsDirty,
  requestMarkdownClose,
} from '../../../markdown/src/main/markdown-main'
import {
  createHtmlPresentView,
  createHtmlView,
  htmlIsDirty,
  requestHtmlClose,
} from '../../../html/src/main/html-main'
import {
  createPdfView,
  clearPdfDirty,
  pdfIsDirty,
  requestPdfClose,
} from '../../../pdf/src/main/pdf-main'
import {
  createSheetsView,
  nudgeQueuedWorkbook,
  queueWorkbookForView,
  requestSheetsClose,
  setActiveSheetsWebContents,
  setSheetsNewBlank,
  sheetsPendingEditCount,
} from '../../../sheets/src/main/sheets-main'
import {
  createSlidesView,
  requestSlidesClose,
  setActiveSlidesWebContents,
  slidesIsDirty,
} from '../../../slides/src/main/slides-main'
import type { DocumentTabKind, OpenDocumentTab, TabKind, TabSummary } from '../shared/tabs-api'

interface TabRecord {
  id: string
  kind: TabKind
  /** null for the Home tab — it's rendered by the shell window's own webContents */
  view: WebContentsView | null
  title: string
  filePath?: string
  /** chrome-free Present tab: no file, no editor menu or save/export targets */
  present?: boolean
}

/** must match the tab strip's rendered height (apps/shell/src/renderer/src/TabBar.tsx) */
const TAB_STRIP_HEIGHT = 40
const HOME_ID = 'home'

/**
 * Owns every open tab (Home + docs + sheets) inside the shell's single
 * BrowserWindow. Docs/sheets tabs are WebContentsView children of that
 * window; only the active one is visible at a time. Home has no view of its
 * own — hiding every other tab reveals the shell window's own content.
 */
export class TabManager {
  private readonly tabs: TabRecord[] = [
    { id: HOME_ID, kind: 'home', view: null, title: 'GenOffice' },
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
  /** Sheets renderer mounted ahead of the next open: parsing its bundle and
   *  booting Univer is the bulk of a workbook's open time, and the shell hands
   *  the path over after mount anyway. */
  private spareSheetsView: WebContentsView | null = null
  private spareSheetsTimer: ReturnType<typeof setTimeout> | null = null

  constructor(
    private readonly shellWindow: BrowserWindow,
    private readonly onChanged: () => void,
    private readonly applyMenuFor: (kind: TabKind) => void,
    /** localized placeholder title for a tab that has no file yet */
    private readonly untitledTitleFor?: (kind: TabKind) => string,
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
    shellWindow.webContents.once('did-finish-load', () => this.scheduleSpareSheetsView(1500))
  }

  private scheduleSpareSheetsView(delayMs: number): void {
    if (process.env.GENOFFICE_NO_SPARE_VIEW || this.spareSheetsTimer) return
    this.spareSheetsTimer = setTimeout(() => {
      this.spareSheetsTimer = null
      if (this.spareSheetsView || this.shellWindow.isDestroyed()) return
      const view = createSheetsView({ includeAiHandlers: false })
      // registering the session made the spare the menu-action target
      const active = this.tabs.find((t) => t.id === this.activeId)
      setActiveSheetsWebContents(
        active?.kind === 'sheets' && active.view ? active.view.webContents : null,
      )
      this.shellWindow.contentView.addChildView(view)
      view.setVisible(false)
      view.setBounds(this.contentBounds())
      view.webContents.once('render-process-gone', () => {
        if (this.spareSheetsView !== view) return
        this.spareSheetsView = null
        view.webContents.close()
      })
      this.spareSheetsView = view
    }, delayMs)
  }

  private takeSpareSheetsView(): WebContentsView | null {
    const view = this.spareSheetsView
    this.spareSheetsView = null
    return view && !view.webContents.isDestroyed() ? view : null
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
    return { x: 0, y: TAB_STRIP_HEIGHT, width, height: Math.max(0, height - TAB_STRIP_HEIGHT) }
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
  private trackHtmlFullScreen(id: string, view: WebContentsView): void {
    view.webContents.on('enter-html-full-screen', () => {
      this.htmlFullScreenId = id
      this.layout()
    })
    view.webContents.on('leave-html-full-screen', () => {
      if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
      this.layout()
    })
  }

  /** re-fit the active tab's view after a window resize */
  layout(): void {
    // Deferred resize layouts can land after the shell window was closed.
    if (this.shellWindow.isDestroyed()) return
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (active?.view) active.view.setBounds(this.contentBounds())
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
    const tabs = this.tabs.filter((tab) => tab.kind !== 'home' && !tab.present && tab.view)
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

  /** unsaved-changes state of one tab, whichever family owns it */
  private async tabIsDirty(tab: TabRecord): Promise<boolean> {
    const wc = tab.view?.webContents
    if (!wc || wc.isDestroyed()) return false
    switch (tab.kind) {
      case 'sheets':
        return sheetsPendingEditCount(wc.id) > 0
      case 'pdf':
        return pdfIsDirty(wc.id)
      case 'markdown':
        return markdownIsDirty(wc.id)
      case 'html':
        return htmlIsDirty(wc.id)
      case 'slides':
        return slidesIsDirty(wc.id)
      case 'docs':
        return docsQueryDirty(wc)
      default:
        return false
    }
  }

  openHomeTab(): void {
    this.activateTab(HOME_ID)
  }

  openDocsTab(
    openPath?: string,
    options?: { newBlank?: boolean; aiContent?: AiDocContent },
  ): string {
    const view = createDocsView(openPath)
    const id = `t${this.nextId++}`
    if (options?.newBlank) markDocsNewBlank(view.webContents.id)
    if (options?.aiContent) queueDocsAiContent(view.webContents.id, options.aiContent)
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'docs',
      view,
      title: openPath ? basename(openPath) : this.untitled('docs', 'GenOffice Docs'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openSheetsTab(openPath?: string, options?: { newBlank?: boolean }): string {
    if (options?.newBlank) setSheetsNewBlank()
    const spare = this.takeSpareSheetsView()
    const view = spare ?? createSheetsView({ includeAiHandlers: false })
    // bind the path to this tab's webContents: a multi-select Open creates
    // several sheets tabs in one loop, so a single global path would be
    // overwritten before the earlier tabs consume it
    if (openPath) {
      queueWorkbookForView(view.webContents, openPath)
      if (spare) nudgeQueuedWorkbook(view.webContents)
    }
    const id = `t${this.nextId++}`
    if (!spare) {
      this.shellWindow.contentView.addChildView(view)
      view.setVisible(false)
    }
    this.trackHtmlFullScreen(id, view)
    this.scheduleSpareSheetsView(3000)
    this.tabs.push({
      id,
      kind: 'sheets',
      view,
      title: openPath ? basename(openPath) : this.untitled('sheets', 'AI Sheets'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openSlidesTab(openPath?: string): string {
    const view = createSlidesView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'slides',
      view,
      title: openPath ? basename(openPath) : this.untitled('slides', 'AI Slides'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openPdfTab(openPath: string): string {
    const view = createPdfView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({ id, kind: 'pdf', view, title: basename(openPath), filePath: openPath })
    this.activateTab(id)
    return id
  }

  /** Remount the tab's renderer so it re-reads its file from disk (View > Reload). */
  reloadTab(id: string): void {
    const tab = this.tabs.find((t) => t.id === id)
    const wc = tab?.view?.webContents
    if (!wc || wc.isDestroyed()) return
    if (tab.kind === 'pdf') clearPdfDirty(wc.id)
    wc.reload()
  }

  openMarkdownTab(openPath?: string): string {
    const view = createMarkdownView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'markdown',
      view,
      title: openPath ? basename(openPath) : this.untitled('markdown', 'AI Markdown'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openHtmlTab(openPath?: string): string {
    const view = createHtmlView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'html',
      view,
      title: openPath ? basename(openPath) : this.untitled('html', 'AI HTML'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  /** Present → New tab: a chrome-free html tab showing the owner tab's live preview */
  openHtmlPresentTab(owner: WebContents, title: string): string {
    const view = createHtmlPresentView(owner, title)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.tabs.push({
      id,
      kind: 'html',
      view,
      title: title || this.untitled('html', 'AI HTML'),
      present: true,
    })
    this.activateTab(id)
    return id
  }

  activateTab(id: string): void {
    const target = this.tabs.find((t) => t.id === id)
    if (!target) return
    for (const t of this.tabs) t.view?.setVisible(t.id === id)
    if (target.view) target.view.setBounds(this.contentBounds())
    this.activeId = id
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
    setActiveDocsResolver(target.kind === 'docs' ? () => target.view!.webContents : () => null)
    if (target.kind === 'sheets' && target.view) setActiveSheetsWebContents(target.view.webContents)
    if (target.kind === 'slides' && target.view) setActiveSlidesWebContents(target.view.webContents)
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
    for (const tab of this.tabs) {
      if (tab.filePath !== oldPath) continue
      tab.filePath = newPath
      tab.title = basename(newPath)
      if (tab.view) affected.push({ kind: tab.kind, webContents: tab.view.webContents })
    }
    if (affected.length > 0) this.onChanged()
    return affected
  }

  /** sheets tabs whose renderer reports unsaved journal edits (shell-close guard) */
  dirtySheetsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter(
        (t) => t.kind === 'sheets' && t.view && sheetsPendingEditCount(t.view.webContents.id) > 0,
      )
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** pdf tabs whose renderer reports unsaved markups/form edits (shell-close guard) */
  dirtyPdfTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'pdf' && t.view && pdfIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** markdown tabs whose renderer reports unsaved edits (shell-close guard) */
  dirtyMarkdownTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'markdown' && t.view && markdownIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** html tabs whose renderer reports unsaved edits (shell-close guard) */
  dirtyHtmlTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'html' && t.view && htmlIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** slides tabs whose main-process session has unsaved edits (shell-close guard) */
  dirtySlidesTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'slides' && t.view && slidesIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** all live docs tabs — dirtiness lives renderer-side, caller queries async (shell-close guard) */
  docsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'docs' && t.view)
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** all live slides tabs (MCP bridge resolves its new tab's webContents through this) */
  slidesTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'slides' && t.view)
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** all live sheets tabs (MCP bridge resolves its new tab's webContents through this) */
  sheetsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'sheets' && t.view)
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** closes whichever tab is currently active; no-op for Home (Cmd+W target) */
  closeActiveTab(): void {
    void this.closeTab(this.activeId)
  }

  async closeTab(id: string): Promise<void> {
    if (id === HOME_ID) return
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab || this.closingIds.has(id)) return
    let closeGuard =
      tab.view &&
      (tab.kind === 'sheets' && sheetsPendingEditCount(tab.view.webContents.id) > 0
        ? requestSheetsClose
        : tab.kind === 'pdf' && pdfIsDirty(tab.view.webContents.id)
          ? requestPdfClose
          : tab.kind === 'markdown' && markdownIsDirty(tab.view.webContents.id)
            ? requestMarkdownClose
            : tab.kind === 'html' && htmlIsDirty(tab.view.webContents.id)
              ? requestHtmlClose
              : tab.kind === 'slides' && slidesIsDirty(tab.view.webContents.id)
                ? requestSlidesClose
                : null)
    // docs dirty state lives in the renderer and needs an async query; skip the guard when clean (avoids a flash activation)
    if (!closeGuard && tab.kind === 'docs' && tab.view) {
      this.closingIds.add(id)
      try {
        if (await docsQueryDirty(tab.view.webContents)) closeGuard = requestDocsClose
      } finally {
        this.closingIds.delete(id)
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
    if (this.activeId === id) {
      const fallback = this.tabs[idx - 1] ?? this.tabs[0]
      this.activateTab(fallback.id)
    } else {
      this.onChanged()
    }
    if (removed.view) {
      removed.view.setVisible(false)
      this.shellWindow.contentView.removeChildView(removed.view)
      if (removed.kind === 'docs') {
        // webContents.close()/.destroy() on a closed docs tab wedges Electron's whole
        // UI thread in a native modal run loop (reproduced consistently; survives
        // close() vs destroy(), teardown ordering, deferring, and disabling
        // accessibility support — looks like an upstream WebContentsView/Chromium
        // issue, not something fixable from here). Detaching without destroying
        // avoids the freeze; the orphaned webContents is reclaimed when the app quits.
        teardownDocsRenderer(removed.view.webContents)
      } else {
        removed.view.webContents.close()
      }
    }
  }

  /** Remove a tab from the strip WITHOUT destroying its WebContentsView and
   *  hand it to the caller ("Open in New Window" — the live document, unsaved
   *  edits included, moves into a detached editor window). Null while a close
   *  prompt is pending on the tab. */
  detachTab(
    id: string,
  ): { view: WebContentsView; kind: TabKind; title: string; filePath?: string } | null {
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

  /** the editor tab showing this file, whichever module owns it (path compared after resolving links) */
  findTabByPath(
    path?: string,
  ): { id: string; kind: TabKind; webContents: WebContents } | undefined {
    const wanted = canonicalPath(path)
    const tab = this.tabs.find(
      (t) => t.view && t.filePath && !t.present && canonicalPath(t.filePath) === wanted,
    )
    return tab?.view ? { id: tab.id, kind: tab.kind, webContents: tab.view.webContents } : undefined
  }

  webContentsForTab(id: string): WebContents | undefined {
    return this.tabs.find((t) => t.id === id)?.view?.webContents
  }

  private findTabOfKindByPath(kind: TabKind, path?: string): string | undefined {
    const wanted = canonicalPath(path)
    return this.tabs.find(
      (t) =>
        t.kind === kind &&
        t.view &&
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

export function canonicalPath(path: string | undefined): string | undefined {
  if (path === undefined) return undefined
  try {
    return realpathSync.native(path)
  } catch {
    return path
  }
}
