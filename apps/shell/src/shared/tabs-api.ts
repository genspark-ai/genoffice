export type TabKind = 'home' | 'docs' | 'sheets' | 'slides' | 'pdf' | 'markdown' | 'html'

/** a tab that holds a document — every kind except the Home screen */
export type DocumentTabKind = Exclude<TabKind, 'home'>

/** one open tab in the top tab strip; Home is always id 'home' and not closable */
export interface TabSummary {
  id: string
  kind: TabKind
  title: string
  closable: boolean
  active: boolean
  /** absolute path behind the tab; absent while the document is untitled */
  filePath?: string
}

/**
 * One document open in an editor tab, as reported to the MCP `open_documents`
 * tool. Unlike `TabSummary` this excludes Home and chrome-free Present tabs (no
 * file, nothing an agent could read or close) and carries the unsaved-changes
 * state, which the shell resolves per family.
 */
export interface OpenDocumentTab {
  id: string
  kind: DocumentTabKind
  title: string
  filePath?: string
  active: boolean
  dirty: boolean
}

export interface TabsApi {
  list(): Promise<TabSummary[]>
  activate(id: string): Promise<void>
  close(id: string): Promise<void>
  /**
   * pop up the native "all tabs" list menu at (x, y) in window CSS coordinates.
   * Native because the content area below the strip is a WebContentsView that
   * would cover any DOM dropdown rendered by the shell.
   */
  showMenu(x: number, y: number): Promise<void>
  /**
   * pop up the native "+" new-file menu (new doc/sheet/slides, open local
   * file) at (x, y) in window CSS coordinates. Native for the same reason
   * as showMenu.
   */
  showNewMenu(x: number, y: number): Promise<void>
  /**
   * pop up the native per-tab context menu (Open in New Window / Close) at
   * (x, y) in window CSS coordinates. Native for the same reason as showMenu.
   */
  showTabMenu(id: string, x: number, y: number): Promise<void>
  /** detach a docs/sheets tab into its own window ("Open in New Window") */
  detach(id: string): Promise<void>
  /**
   * pop up the application menu (File / Edit / View …) at (x, y). Windows and
   * Linux hide the native menu bar under the tab strip; macOS keeps the
   * system menu bar and never shows the button.
   */
  showAppMenu(x: number, y: number): Promise<void>
  /** move a tab to a new index in the strip; Home stays pinned at index 0 */
  reorder(id: string, toIndex: number): Promise<void>
  /** subscribe to tab list changes (open/close/activate/title updates); returns unsubscribe */
  onChanged(handler: (tabs: TabSummary[]) => void): () => void
  /**
   * fire-and-forget: a pointerdown landed on the shell chrome (tab strip).
   * Document tabs are sibling WebContentsViews that see neither the event nor
   * a focus change, so the shell relays it for them to dismiss popovers.
   */
  notifyChromePressed(): void
  /** the main-process broadcast that notifyChromePressed (and a window drag)
   * triggers; the shell's own popovers subscribe so a title-bar drag — which
   * produces no DOM event — still dismisses them */
  onChromePressed(handler: () => void): () => void
}

export const TABS_CHANNELS = {
  list: 'tabs:list',
  activate: 'tabs:activate',
  close: 'tabs:close',
  showMenu: 'tabs:show-menu',
  showNewMenu: 'tabs:show-new-menu',
  showTabMenu: 'tabs:show-tab-menu',
  detach: 'tabs:detach',
  showAppMenu: 'tabs:show-app-menu',
  reorder: 'tabs:reorder',
  changed: 'tabs:changed',
  chromePressed: 'tabs:chrome-pressed',
} as const
