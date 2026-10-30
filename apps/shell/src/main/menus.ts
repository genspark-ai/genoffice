import { readFileSync } from 'node:fs'
import { app, BrowserWindow, dialog, Menu, nativeImage } from 'electron'
import type { MenuItemConstructorOptions, NativeImage, WebContents } from 'electron'
import { showOpenDialogWithMemory } from '@genoffice/electron-utils'
import { changeUiScale } from './ui-scale'
import { TabManager } from './tab-manager'
import { focusedDetachedTab, isDetachedEditorWindow } from './detached-windows'
import {
  setDocsExtraFileMenuItems,
  setSheetsExtraFileMenuItems,
  setSlidesExtraFileMenuItems,
} from './app-registry'
import menuDocxIcon1x from './assets/menu-docx.png?asset'
import menuDocxIcon2x from './assets/menu-docx@2x.png?asset'
import menuXlsxIcon1x from './assets/menu-xlsx.png?asset'
import menuXlsxIcon2x from './assets/menu-xlsx@2x.png?asset'
import menuPptxIcon1x from './assets/menu-pptx.png?asset'
import menuPptxIcon2x from './assets/menu-pptx@2x.png?asset'
import menuPdfIcon1x from './assets/menu-pdf.png?asset'
import menuPdfIcon2x from './assets/menu-pdf@2x.png?asset'
import menuMdIcon1x from './assets/menu-md.png?asset'
import menuMdIcon2x from './assets/menu-md@2x.png?asset'
import menuHtmlIcon1x from './assets/menu-html.png?asset'
import menuHtmlIcon2x from './assets/menu-html@2x.png?asset'
import menuHomeIcon1x from './assets/menu-home.png?asset'
import menuHomeIcon2x from './assets/menu-home@2x.png?asset'
import type { TabKind } from '../shared/tabs-api'

/**
 * The application-menu machinery the shell builds around its own and the
 * editor apps' menus: the shared icons, the menu injections every tab kind
 * gets (user guide, View scale items, back-to-home), the macOS dock menu,
 * and the focused-target resolution the pdf/markdown/html menu commands use.
 * The four whole-menu builders (home/pdf/markdown/html) stay in index.ts —
 * they close over the tab/new-file wiring and the locale table there — and
 * call into this module's pieces.
 */

/** the locale keys this module resolves through the host (index's tm) */
export type MenuStringKey =
  | 'menuUserGuide'
  | 'menuHelp'
  | 'menuView'
  | 'menuUiZoomIn'
  | 'menuUiZoomOut'
  | 'menuUiZoomReset'
  | 'backToHome'
  | 'filterSupported'
  | 'menuHome'
  | 'menuNewDoc'
  | 'menuNewSheet'
  | 'menuNewSlide'
  | 'menuNewMarkdown'
  | 'menuNewHtml'
  | 'menuNewPdf'
  | 'menuOpen'

export interface MenuHost {
  translate: (key: MenuStringKey) => string
  appSettingsPath: () => string
  getTabManager: () => TabManager | null
  getShellWindow: () => BrowserWindow | null
  openDocumentPath: (filePath: string) => boolean
  /** every extension the open dialog offers, mirrored from the app registry */
  openDialogExtensions: readonly string[]
  /** the new-file tab creators the dock menu offers */
  newDocTab: () => void
  newSheetTab: () => void | Promise<void>
  newSlideTab: () => void
  newMarkdownTab: () => void
  newHtmlTab: () => void
  newPdfTab: () => void | Promise<void>
}

// electron-vite emits ?asset files under hashed names, which breaks nativeImage's
// automatic `@2x` sibling lookup — attach the retina representation by hand
export function loadMenuIcon(path1x: string, path2x: string): NativeImage {
  const icon = nativeImage.createFromPath(path1x)
  icon.addRepresentation({ scaleFactor: 2, buffer: readFileSync(path2x) })
  return icon
}

// loaded once, not on every menu open
export interface MenuIconSet {
  docx: NativeImage
  xlsx: NativeImage
  pptx: NativeImage
  pdf: NativeImage
  md: NativeImage
  html: NativeImage
  home: NativeImage
}
let menuIconCache: MenuIconSet | null = null
export function menuIcons(): MenuIconSet {
  menuIconCache ??= {
    docx: loadMenuIcon(menuDocxIcon1x, menuDocxIcon2x),
    xlsx: loadMenuIcon(menuXlsxIcon1x, menuXlsxIcon2x),
    pptx: loadMenuIcon(menuPptxIcon1x, menuPptxIcon2x),
    pdf: loadMenuIcon(menuPdfIcon1x, menuPdfIcon2x),
    md: loadMenuIcon(menuMdIcon1x, menuMdIcon2x),
    html: loadMenuIcon(menuHtmlIcon1x, menuHtmlIcon2x),
    home: loadMenuIcon(menuHomeIcon1x, menuHomeIcon2x),
  }
  return menuIconCache
}

export const TAB_MENU_ICON: Record<TabKind, keyof MenuIconSet> = {
  home: 'home',
  docs: 'docx',
  sheets: 'xlsx',
  slides: 'pptx',
  pdf: 'pdf',
  markdown: 'md',
  html: 'html',
  help: 'home',
}

/**
 * Put the manual in whatever menu is currently installed.
 *
 * Every tab kind builds its own application menu, which is how File and Edit
 * stay tab-shaped — and which is also why an item added to one builder is
 * missing from the other six. F1 worked from Home and nowhere else, and the
 * manual's own shortcut table was wrong for six of the seven tab kinds.
 *
 * Injecting after the builder has run is the one place that cannot be
 * forgotten: a new tab kind gets the manual for free, and a builder that
 * already has it is left alone rather than getting a second copy.
 */
export function withUserGuide(host: MenuHost): void {
  const menu = Menu.getApplicationMenu()
  if (!menu) return
  const userGuide = {
    label: host.translate('menuUserGuide'),
    accelerator: 'F1',
    click: () => host.getTabManager()?.openHelpTab(),
  }
  const help = menu.items.find((item) => item.role === 'help')
  // Slides and Sheets build a File/Edit/View template with no Help menu at
  // all, so there is nothing to insert into — one is created instead. Without
  // this the manual would be reachable from four of the seven tab kinds.
  if (!help) {
    menu.append(
      Menu.buildFromTemplate([
        { role: 'help', label: host.translate('menuHelp'), submenu: [userGuide] },
      ]).items[0]!,
    )
    return
  }
  const submenu = help.submenu
  if (!submenu || submenu.items.some((i) => i.accelerator === 'F1')) return
  // A separator survives a template only *between* two real items — Electron
  // drops one at either end, so `[userGuide, separator]` builds a single item
  // and the destructure below used to hand `undefined` to `insert`, which threw
  // "Invalid item" and took the main process down on the next tab switch. The
  // third entry is filler, discarded; the type check is what stops this from
  // silently regressing if Electron's rule ever changes again.
  const [item, separator] = Menu.buildFromTemplate([
    userGuide,
    { type: 'separator' },
    { role: 'undo' },
  ]).items
  if (!item) return
  submenu.insert(0, item)
  if (separator?.type === 'separator') submenu.insert(1, separator)
}

/**
 * The View › interface-scale items, shared by the home menu and every editor
 * tab. CmdOrCtrl+Alt (not Shift): Shift+= is Sheets insert-cells and Docs
 * superscript, and the View menu now exists on editor tabs too, where those
 * accelerators are taken.
 */
export function uiScaleMenuItems(host: MenuHost): MenuItemConstructorOptions[] {
  return [
    {
      label: host.translate('menuUiZoomIn'),
      accelerator: 'CmdOrCtrl+Alt+=',
      click: () => changeUiScale(host.appSettingsPath(), 1),
    },
    {
      label: host.translate('menuUiZoomOut'),
      accelerator: 'CmdOrCtrl+Alt+-',
      click: () => changeUiScale(host.appSettingsPath(), -1),
    },
    { type: 'separator' },
    {
      label: host.translate('menuUiZoomReset'),
      accelerator: 'CmdOrCtrl+Alt+0',
      click: () => changeUiScale(host.appSettingsPath(), 0),
    },
  ]
}

export function viewMenuTemplate(host: MenuHost): MenuItemConstructorOptions {
  return { label: host.translate('menuView'), submenu: uiScaleMenuItems(host) }
}

/**
 * Interface-scale items in the View menu of every tab kind. Editor menus are
 * built in the editor apps (docs/sheets/slides) or by the three shell-owned
 * builders, so rather than teaching seven menu builders about the scale, the
 * current menu is post-processed here — the same shape as withUserGuide below.
 * An existing View menu (docs, slides: document zoom) gains the items after a
 * separator; the rest get a fresh View menu inserted before the Help menu that
 * withUserGuide has just guaranteed.
 */
export function withViewMenu(host: MenuHost): void {
  const menu = Menu.getApplicationMenu()
  if (!menu) return
  const hasScale = menu.items.some((item) =>
    item.submenu?.items.some((i) => i.accelerator === 'CmdOrCtrl+Alt+='),
  )
  if (hasScale) return
  const view = menu.items.find((item) => item.label === host.translate('menuView'))
  if (view?.submenu) {
    // build the separator together with the items: Electron drops a separator
    // that sits alone in a template
    for (const item of Menu.buildFromTemplate([{ type: 'separator' }, ...uiScaleMenuItems(host)])
      .items) {
      view.submenu.append(item)
    }
    return
  }
  const viewItem = Menu.buildFromTemplate([viewMenuTemplate(host)]).items[0]!
  const helpIdx = menu.items.findIndex((item) => item.role === 'help')
  if (helpIdx === -1) menu.append(viewItem)
  else menu.insert(helpIdx, viewItem)
}

/** Menu-command target for the editors whose menus resolve through the shell
 *  (pdf / markdown / html): the focused detached window of that kind first,
 *  else the shell's active tab of that kind. */
export function activeEditorTarget(
  host: MenuHost,
  kind: TabKind,
  fromTabs: () => { id: string; webContents: WebContents; filePath?: string } | undefined,
): { id: string; webContents: WebContents; filePath?: string } | undefined {
  const detachedTab = focusedDetachedTab()
  if (detachedTab?.kind === kind) return detachedTab
  return fromTabs()
}

/** ⌘W / File > Close targets the focused window: a detached editor closes
 *  itself (its own unsaved-changes guard runs), otherwise the shell's active tab
 *  — the same rule the docs / sheets / slides close hooks follow */
export function closeFocusedEditor(host: MenuHost): void {
  const focused = BrowserWindow.getFocusedWindow()
  if (focused && focused !== host.getShellWindow() && isDetachedEditorWindow(focused))
    focused.close()
  else host.getTabManager()?.closeActiveTab()
}

/** the open dialog the File > Open menu item and the home menu share */
export async function openFileViaDialog(host: MenuHost): Promise<void> {
  const win = host.getShellWindow() ?? BrowserWindow.getFocusedWindow()
  if (!win) return
  const result = await showOpenDialogWithMemory(dialog, win, {
    filters: [
      { name: host.translate('filterSupported'), extensions: [...host.openDialogExtensions] },
    ],
    properties: ['openFile', 'multiSelections'],
  })
  if (!result.canceled) for (const path of result.filePaths) host.openDocumentPath(path)
}

/** every module's File menu gets a way back to the launcher */
export function installBackToHomeItems(host: MenuHost): void {
  const backToHomeItem: MenuItemConstructorOptions = {
    label: host.translate('backToHome'),
    accelerator: 'Shift+CmdOrCtrl+H',
    click: () => host.getTabManager()?.openHomeTab(),
  }
  setDocsExtraFileMenuItems([backToHomeItem])
  setSheetsExtraFileMenuItems([backToHomeItem])
  setSlidesExtraFileMenuItems([backToHomeItem])
}

export function installDockMenu(host: MenuHost): void {
  if (process.platform !== 'darwin') return
  app.dock?.setMenu(
    Menu.buildFromTemplate([
      { label: host.translate('menuHome'), click: () => host.getTabManager()?.openHomeTab() },
      {
        label: host.translate('menuNewDoc'),
        click: () => host.newDocTab(),
      },
      {
        label: host.translate('menuNewSheet'),
        click: () => void host.newSheetTab(),
      },
      { label: host.translate('menuNewSlide'), click: () => host.newSlideTab() },
      { label: host.translate('menuNewMarkdown'), click: () => host.newMarkdownTab() },
      { label: host.translate('menuNewPdf'), click: () => void host.newPdfTab() },
    ]),
  )
}
