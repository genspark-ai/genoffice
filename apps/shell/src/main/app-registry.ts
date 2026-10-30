import type { BrowserWindow, WebContents, WebContentsView } from 'electron'

import {
  buildDocsMenu,
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
  createSheetsView,
  installSheetsMenu,
  nudgeQueuedWorkbook,
  queueWorkbookForView,
  requestSheetsClose,
  setActiveSheetsWebContents,
  setSheetsNewBlank,
  sheetsPendingEditCount,
} from '../../../sheets/src/main/sheets-main'
import {
  createSlidesView,
  discardSlidesRecovery,
  installSlidesMenu,
  requestSlidesClose,
  setActiveSlidesWebContents,
  slidesIsDirty,
} from '../../../slides/src/main/slides-main'
import {
  clearPdfDirty,
  createPdfView,
  pdfIsDirty,
  requestPdfClose,
} from '../../../pdf/src/main/pdf-main'
import {
  createMarkdownView,
  markdownDiscardPendingAssets,
  markdownIsDirty,
  markdownReadText,
  markdownSaveToPath,
  requestMarkdownClose,
} from '../../../markdown/src/main/markdown-main'
import {
  createHtmlPresentView,
  createHtmlView,
  htmlDiscardPendingAssets,
  htmlIsDirty,
  htmlReadText,
  htmlSaveToPath,
  requestHtmlClose,
} from '../../../html/src/main/html-main'
import type { DocumentTabKind, TabKind } from '../shared/tabs-api'
import type { EditorFamily } from './mcp/tools/formats'
import type { DocsControl } from './mcp/tools/document-tools'
import type { SheetsControl } from './mcp/tools/sheets-tools'
import type { SlidesControl } from './mcp/tools/slides-tools'

// ── the AppDescriptor registry ────────────────────────────────────────────
//
// One row per bundled document app: everything the shell needs to host the
// app's documents — the view factory, dirty/close guards, routing extensions,
// untitled title, menu builder and the MCP agent surface — declared once so
// app-routing, TabManager, the close guards, detached windows and the MCP
// bridge all look up this table instead of switching on families.
//
// The descriptors are declared centrally here rather than as a `descriptor()`
// export on each app's main module: the six apps today are plain modules with
// no shell-facing contract, several descriptor faces (the shell-owned pdf/
// markdown/html menus, the runtime MCP controls) only exist on the shell side,
// and a central table keeps the wiring diff to the app modules at zero. When
// third-party apps arrive, this table is the seam they register through.
//
// This file is also the shell main process's single static-import site for the
// six app main modules (the re-exports below): "which app code does the shell
// link against" has one answer, and swapping the static imports for lazy ones
// later touches exactly one file.

/** what the MCP `open_documents` tool can do with one open document of an app */
export interface AgentDocumentIo {
  /** read the document's live content for the agent */
  read?: (wc: WebContents) => Promise<unknown>
  /** write the document to a path (open_documents save / save-close) */
  save?: (wc: WebContents, filePath: string) => Promise<unknown>
  /** release staged assets a discard leaves behind (markdown/html asset files, slides recovery copies) */
  discard?: (wc: WebContents) => void | Promise<void>
  /** assets staged beside the document file only: an untitled tab has none to release */
  discardNeedsFilePath?: boolean
  /** the family cannot be saved through the tool (the pdf app is a viewer); the message the save surfaces */
  saveDenied?: string
  /** refusal when a *dirty* tab of this family is asked to save-close (pdf's longer instruction) */
  saveCloseDenied?: (title: string) => string
}

/** app-specific steps around opening a tab of this app */
export interface AppOpenHooks {
  /** app-global state to set before the view exists (sheets: the in-memory blank flag) */
  prepare?: (options: { newBlank?: boolean }) => void
  /** per-view markers to apply right after creation (docs: new-blank and queued AI content) */
  afterCreate?: (wcId: number, options: { newBlank?: boolean; aiContent?: AiDocContent }) => void
  /** the app opens measurably faster from a pre-warmed spare view (sheets: renderer boot dominates) */
  spare?: {
    /** hand a just-adopted spare the pending file (sheets: queued workbook + renderer nudge) */
    adopt?: (view: WebContentsView, openPath: string) => void
  }
}

/** untitled-title keys the shell's locale tables carry, as a literal union so a
 *  key the tables do not define fails typecheck at the localized lookup */
export type UntitledTitleKey =
  | 'untitledDoc'
  | 'untitledSheet'
  | 'untitledDeck'
  | 'untitledMarkdown'
  | 'untitledHtml'
  | 'untitledPdf'

export interface AppDescriptor {
  /** registry id: the tab kind the shell uses for this app's documents */
  id: DocumentTabKind
  /** the MCP layer's editor family for this app */
  family: EditorFamily
  /** extensions the open route files into this app (no leading dots, lowercase) */
  extensions: readonly string[]
  /** tab title before the first save, when no localized title provider is wired */
  untitledTitle: string
  /** i18n key for the localized untitled title (the shell's locale tables) */
  untitledTitleKey: UntitledTitleKey
  /** build the WebContentsView for a tab of this app; a path opens the file, none opens a blank document */
  /** build the WebContentsView for a tab of this app; a path opens the file, none opens a blank document */
  createView: (openPath?: string) => WebContentsView
  /** synchronous unsaved-changes answer; the five families that keep dirty state in the main process */
  isDirty?: (wc: WebContents) => boolean
  /** async unsaved-changes answer (docs keeps its dirty state renderer-side) */
  queryDirty?: (wc: WebContents) => Promise<boolean>
  /** run the save/don't-save/cancel prompt; true when the tab may close */
  requestClose: (wc: WebContents, win: BrowserWindow) => Promise<boolean>
  /** editor teardown beyond `view.webContents.close()` (docs: main-process bookkeeping) */
  teardown?: (wc: WebContents) => void
  /** clear the dirty flag before a reload (pdf: pending markup edits) */
  clearDirty?: (wc: WebContents) => void
  /** app-owned menu builder for tabs of this kind; shell-owned builders are registered at boot */
  installMenu?: () => void
  /** app-specific steps around opening a tab */
  open?: AppOpenHooks
  /** chrome-free present view of another tab's document (html only today) */
  createPresentView?: (owner: WebContents, title: string) => WebContentsView
  /**
   * Claim the process-global "active editor of this app" target while one of
   * this app's tabs is focused (docs/sheets/slides each keep one; the getter
   * resolves lazily so a re-created view is picked up, and null releases it).
   */
  claimActiveTarget?: (getActive: () => WebContents | null) => void
  /** what the MCP open-documents tool can read/save/discard for this app */
  agentControl?: AgentDocumentIo
  /**
   * This app ships its own plain Ctrl/Cmd+K binding (docs: Word's hyperlink
   * dialog) and keeps it: the shell's palette summon yields inside it — the
   * palette stays reachable through the global hotkey and the tray.
   */
  ownsPlainCtrlK?: boolean
}

/** the requestClose signature, for close-guard plumbing */
export type AppCloseGuard = AppDescriptor['requestClose']

/** the six bundled apps, in registry/routing order (first match wins) */
export const documentApps: readonly AppDescriptor[] = [
  {
    id: 'docs',
    family: 'docx',
    extensions: ['docx', 'odt', 'ott'],
    untitledTitle: 'GenOffice Docs',
    untitledTitleKey: 'untitledDoc',
    createView: createDocsView,
    queryDirty: docsQueryDirty,
    requestClose: requestDocsClose,
    teardown: teardownDocsRenderer,
    installMenu: buildDocsMenu,
    // Ctrl/Cmd+K is Word's insert-hyperlink dialog, and the shell's palette
    // summon yields to it (the palette keeps the global hotkey and the tray)
    ownsPlainCtrlK: true,
    open: {
      afterCreate: (wcId, { newBlank, aiContent }) => {
        if (newBlank) markDocsNewBlank(wcId)
        if (aiContent) queueDocsAiContent(wcId, aiContent)
      },
    },
    claimActiveTarget: (getActive) => setActiveDocsResolver(getActive),
  },
  {
    id: 'sheets',
    family: 'xlsx',
    extensions: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'ods', 'ots'],
    untitledTitle: 'AI Sheets',
    untitledTitleKey: 'untitledSheet',
    createView: (openPath) => {
      const view = createSheetsView({
        includeAiHandlers: false,
        openingWorkbook: Boolean(openPath),
      })
      // bind the path to this tab's webContents: a multi-select Open creates
      // several sheets tabs in one loop, so a single global path would be
      // overwritten before the earlier tabs consume it
      if (openPath) queueWorkbookForView(view.webContents, openPath)
      return view
    },
    isDirty: (wc) => sheetsPendingEditCount(wc.id) > 0,
    requestClose: requestSheetsClose,
    installMenu: installSheetsMenu,
    open: {
      prepare: ({ newBlank }) => {
        if (newBlank) setSheetsNewBlank()
      },
      spare: {
        adopt: (view, openPath) => {
          queueWorkbookForView(view.webContents, openPath)
          nudgeQueuedWorkbook(view.webContents)
        },
      },
    },
    claimActiveTarget: (getActive) => {
      const wc = getActive()
      if (wc) setActiveSheetsWebContents(wc)
    },
  },
  {
    id: 'slides',
    family: 'pptx',
    extensions: ['pptx', 'odp', 'otp'],
    untitledTitle: 'AI Slides',
    untitledTitleKey: 'untitledDeck',
    createView: createSlidesView,
    isDirty: (wc) => slidesIsDirty(wc.id),
    requestClose: requestSlidesClose,
    installMenu: installSlidesMenu,
    claimActiveTarget: (getActive) => {
      const wc = getActive()
      if (wc) setActiveSlidesWebContents(wc)
    },
  },
  {
    id: 'pdf',
    family: 'pdf',
    extensions: ['pdf'],
    // openPdfTab always has a path, so the untitled title is never shown; the
    // key exists for the New PDF flow's suggested file name
    untitledTitle: 'GenOffice PDF',
    untitledTitleKey: 'untitledPdf',
    createView: createPdfView,
    isDirty: (wc) => pdfIsDirty(wc.id),
    requestClose: requestPdfClose,
    clearDirty: (wc) => clearPdfDirty(wc.id),
    agentControl: {
      saveDenied: 'a PDF cannot be saved through this tool: the pdf app is a viewer',
      saveCloseDenied: (title) =>
        `"${title}" has unsaved changes, and a PDF cannot be saved through this tool: ` +
        'close it with unsaved:"discard" to drop them, or save it in GenOffice first',
    },
  },
  {
    id: 'markdown',
    family: 'md',
    extensions: ['md', 'markdown', 'txt', 'json'],
    untitledTitle: 'AI Markdown',
    untitledTitleKey: 'untitledMarkdown',
    createView: createMarkdownView,
    isDirty: (wc) => markdownIsDirty(wc.id),
    requestClose: requestMarkdownClose,
    agentControl: {
      read: markdownReadText,
      save: markdownSaveToPath,
      discard: markdownDiscardPendingAssets,
      discardNeedsFilePath: true,
    },
  },
  {
    id: 'html',
    family: 'html',
    extensions: ['html', 'htm'],
    untitledTitle: 'AI HTML',
    untitledTitleKey: 'untitledHtml',
    createView: createHtmlView,
    createPresentView: createHtmlPresentView,
    isDirty: (wc) => htmlIsDirty(wc.id),
    requestClose: requestHtmlClose,
    agentControl: {
      read: htmlReadText,
      save: htmlSaveToPath,
      discard: htmlDiscardPendingAssets,
      discardNeedsFilePath: true,
    },
  },
]

// ── lookups ───────────────────────────────────────────────────────────────

const BY_KIND = new Map(documentApps.map((app) => [app.id, app]))

/** the registered app for a tab kind; undefined for `home` and `help` */
export function appForKind(kind: TabKind): AppDescriptor | undefined {
  return BY_KIND.get(kind as DocumentTabKind)
}

/** the registered app that owns files with this path's extension; undefined when none does */
export function appForPath(path: string): DocumentTabKind | undefined {
  const ext = path.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1]
  if (ext === undefined) return undefined
  return documentApps.find((app) => app.extensions.includes(ext))?.id
}

/** unsaved-changes state of one open view, whichever family owns it: the sync
 *  answer for five families, a renderer round-trip for docs */
export async function appIsDirty(kind: TabKind, wc: WebContents): Promise<boolean> {
  const descriptor = appForKind(kind)
  if (!descriptor) return false
  return descriptor.queryDirty ? descriptor.queryDirty(wc) : (descriptor.isDirty?.(wc) ?? false)
}

// ── menus ─────────────────────────────────────────────────────────────────
//
// docs/sheets/slides build their own editor menus, so their builders live in
// the descriptors above. The pdf/markdown/html menus are chrome the *shell*
// builds around those tabs (their apps have no menu of their own), and the
// builders close over shell state, so index.ts registers them here at boot
// rather than the descriptors importing index.ts back (a cycle).

/** shell-owned menu builders, keyed by tab kind; registered by index.ts at boot */
const shellMenuBuilders = new Map<DocumentTabKind, () => void>()

export function setShellMenuBuilders(builders: {
  [K in DocumentTabKind]?: (() => void) | undefined
}): void {
  for (const [kind, builder] of Object.entries(builders)) {
    if (builder) shellMenuBuilders.set(kind as DocumentTabKind, builder)
    else shellMenuBuilders.delete(kind as DocumentTabKind)
  }
}

/** drop every shell-registered menu builder (tests; future shell teardown) */
export function clearShellMenuBuilders(): void {
  shellMenuBuilders.clear()
}

/**
 * Install the application menu for a tab kind: the app's own builder, or the
 * shell-registered one. Returns false when no app owns `kind` (home, help), so
 * the caller can fall back to the home menu.
 */
export function installAppMenu(kind: TabKind): boolean {
  const descriptor = appForKind(kind)
  if (!descriptor) return false
  const builder = descriptor.installMenu ?? shellMenuBuilders.get(descriptor.id)
  if (!builder) return false
  builder()
  return true
}

// ── the MCP agent surface ─────────────────────────────────────────────────
//
// markdown/html/pdf are driven by app-module functions directly (above);
// docs/sheets/slides need the MCP bridge controls the shell constructs with
// runtime deps, so their AgentDocumentIo is assembled here once the shell has
// the controls (registryAgentIo).

/** the static per-app agent surfaces (markdown, html, pdf); absent for the session families */
export function agentIoFor(kind: DocumentTabKind): AgentDocumentIo | undefined {
  return appForKind(kind)?.agentControl
}

/** the MCP bridge controls for the three session families, as the shell constructs them */
export interface SessionControls {
  docs?: DocsControl
  sheets?: SheetsControl
  slides?: SlidesControl
}

/**
 * The full per-family agent surface for the open-documents tool: the shell's
 * session controls wrapped in the uniform AgentDocumentIo shape, plus every
 * app's static surface. Pass to `createOpenDocumentsControl` as `agents`.
 */
export function registryAgentIo(
  controls: SessionControls = {},
): Partial<Record<DocumentTabKind, AgentDocumentIo>> {
  const agents: Partial<Record<DocumentTabKind, AgentDocumentIo>> = {
    markdown: agentIoFor('markdown'),
    html: agentIoFor('html'),
    pdf: agentIoFor('pdf'),
  }
  if (controls.docs) {
    const docs = controls.docs
    agents.docs = {
      read: (wc) => docs.runCommand(wc.id, 'read_document', {}),
      save: (wc, filePath) =>
        docs.runCommand(wc.id, 'save_document', { path: filePath, overwrite: true }),
    }
  }
  if (controls.sheets) {
    const sheets = controls.sheets
    agents.sheets = {
      read: (wc) => sheets.runCommand(wc.id, 'read_sheet', {}),
      save: (wc, filePath) =>
        sheets.runCommand(wc.id, 'save_sheet', { path: filePath, overwrite: true }),
    }
  }
  if (controls.slides) {
    const slides = controls.slides
    agents.slides = {
      read: (wc) => slides.readDeck(wc.id),
      save: (wc, filePath) => slides.saveDeck(wc.id, filePath, true),
      // slides keeps its own crash-recovery copies; dropping the edits must drop
      // them too, or the next open offers them back
      discard: (wc) => discardSlidesRecovery(wc),
    }
  }
  return agents
}

// ── the shell's single import site for the six app main modules ───────────
// Re-exported for index.ts and the shell modules below it; the names are the
// exact surface the shell consumes, so adding a consumer means adding a name
// here rather than a new direct app import.

export {
  authorizeMcpDocWrite,
  configureDocsRuntime,
  createAiDocument,
  defaultSaveDir,
  docsFileRenamed,
  exportDocsHeadless,
  projectFileRenamed,
  projectFilePaths,
  readRecentFiles,
  readStarredFiles,
  readStarredGroupMap,
  readStarredGroups,
  recordRecentFile,
  registerAiIpc,
  registerDocsIpc,
  registerProjectIpc,
  removeRecentFiles,
  removeStarredFiles,
  replaceRecentFile,
  setActiveDocsResolver,
  setDocsExtraFileMenuItems,
  setDocsFileOpenedHook,
  setDocsFileSavedHook,
  setDocsHostWindowHook,
  setDocsMenuGate,
  setDocsShellHooks,
  setDocsShellWindow,
  setSessionPathResolver,
  setStarredGroup,
  teardownDocsRenderer,
  toggleStarredFile,
  uniquePathIn,
} from '../../../docs/src/main/docs-main'

export {
  authorizeMcpSheetWrite,
  configureSheetsRuntime,
  exportSheetsPdfHeadless,
  hasActiveQueuedWorkbook,
  markSheetsShuttingDown,
  markSheetsUntitledPath,
  markSheetsUnsavedNew,
  pendingUnsavedNewRecoveries,
  resetSheetsShuttingDown,
  resolveSheetsSessionPath,
  sendSheetsMenuAction,
  setActiveSheetsWebContents,
  setSheetsCloseTabHook,
  setSheetsExtraFileMenuItems,
  setSheetsHostWindowHook,
  setSheetsShellWindow,
  setSheetsWorkbookOpenedHook,
  sheetsFileRenamed,
  sheetsSuggestedPathTaken,
  startSheetsCaptureServer,
  stopSheetsSidecar,
  unmarkSheetsUnsavedNew,
} from '../../../sheets/src/main/sheets-main'

export {
  configureSlidesRuntime,
  exportSlidesPdfHeadless,
  readSlidesRecentFiles,
  replaceSlidesRecentFile,
  setSlidesCloseTabHook,
  setSlidesExtraFileMenuItems,
  setSlidesHostWindowHook,
  setSlidesOpenedHook,
  setSlidesShellWindow,
  setSlidesShowBleed,
  slidesFileRenamed,
} from '../../../slides/src/main/slides-main'

export {
  configurePdfRuntime,
  flushPdfSave,
  markPdfUntitledPath,
  pdfFileRenamed,
  pdfIsDirty,
  requestPdfSaveAs,
  sendPdfPrintRequest,
  setPdfRedactionSavedHook,
  setPdfRenamedHook,
  setPdfSaveAsInFlight,
} from '../../../pdf/src/main/pdf-main'
export { blankPdfBuffer } from '../../../pdf/src/main/blank-pdf'
export { PDF_CHANNELS } from '../../../pdf/src/shared/ipc'

export {
  configureMarkdownRuntime,
  exportMarkdownPdfHeadless,
  markdownFileRenamed,
  requestMarkdownSave,
  sendMarkdownExportRequest,
  sendMarkdownPrintRequest,
  setMarkdownDocxExportedHook,
  setMarkdownFileSavedHook,
} from '../../../markdown/src/main/markdown-main'

export {
  configureHtmlRuntime,
  exportHtmlHeadless,
  htmlFileRenamed,
  registerPrivilegedSchemes,
  requestHtmlSave,
  sendHtmlExportRequest,
  sendHtmlPrintRequest,
  setHtmlDocxExportedHook,
  setHtmlDocxExportPrepareHook,
  setHtmlFileSavedHook,
  setHtmlPresentHooks,
  setHtmlProvisionalTitleHook,
} from '../../../html/src/main/html-main'
