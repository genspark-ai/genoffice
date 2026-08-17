/**
 * GenOffice Slides main process — pptx parsing/render-tree building/edit application/saving all live
 * here (Node side). The renderer only gets plain-data RenderSlide; edit intents are sent back
 * here to apply. Structure mirrors apps/docs: exports embeddable configure/register/start for
 * future shell reuse.
 */
import {
  clipboard,
  app,
  BrowserWindow,
  desktopCapturer,
  dialog,
  ipcMain,
  Menu,
  nativeImage,
  session as electronSession,
  shell,
  WebContentsView,
} from 'electron'
import type { WebContents } from 'electron'
import { execFile } from 'node:child_process'
import { readFile, writeFile, rm, stat, mkdir, open } from 'node:fs/promises'
import { createHash, randomUUID } from 'node:crypto'
import { existsSync, mkdirSync } from 'node:fs'
import { userInfo } from 'node:os'
import { dirname, join } from 'node:path'
import { gskApiKey, gskSlideGenerate, setGskProxyUrl } from '@genoffice/ai-search'
import {
  appMenuLabels,
  configuredDefaultSaveDir,
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  toggleDevToolsItem,
} from '@genoffice/electron-utils'
import { getUiLang, normalizeLang, setUiLang } from '@genoffice/i18n'
import { ProjectStore } from '@genoffice/project-store'
import {
  addChart,
  addElement,
  addMedia,
  addModel3d,
  addPicture,
  addSlideComment,
  addSmartArt,
  applyHeaderFooter,
  applyThemeToArchive,
  remapDeckColors,
  addTable,
  copyElementData,
  editPictureSrcRect,
  groupElements,
  ungroupElement,
  findGroupChild,
  editGroupChildTransform,
  patchGroupChildText,
  setGroupChildFont,
  editGroupChildFill,
  editGroupChildStroke,
  patchBodyPrAutofit,
  getElementLink,
  getSlideLinks,
  getRunLinks,
  ensureRunLinkRels,
  readHeaderFooter,
  setElementLink,
  createBlankPptx,
  deleteElement,
  deleteSlide,
  deleteSlideComment,
  duplicateSlide,
  copySlide,
  pasteSlide,
  type SlideBundle,
  insertBlankSlide,
  insertSlideWithLayout,
  listSlideLayouts,
  editTableCellText,
  editTableStructure,
  editTableStyle,
  ensureTableStylePart,
  editChartElement,
  markChartEditable,
  getChartElementData,
  materializeSlide,
  listMasterParts,
  parseMasterPart,
  patchSlideXml,
  TABLE_STYLE_PRESETS,
  setTableColWidth,
  type TableStructureOp,
  type TableStyleEdit,
  EMU_PER_PT,
  getSlideComments,
  getSlideNotes,
  getSlideTransition,
  elementSpid,
  getSlideAnimations,
  setSlideAnimations,
  type SlideAnimation,
  openPptx,
  mergeSlideFromPptx,
  promoteSlideBackground,
  parseTheme,
  pasteElements,
  reorderElement,
  reparseDeck,
  replacePictureBytes,
  savePptx,
  savePptxToFile,
  commitSaved,
  setElementFont,
  replaceAllInDeck,
  mergeTableCells,
  setSlideLayout,
  resetSlideLayout,
  builtinLayoutInfos,
  ensureBuiltinLayout,
  shouldOfferBuiltinLayouts,
  BUILTIN_LAYOUT_PREFIX,
  setSlideSize,
  setPictureOpacity,
  setElementConnection,
  updateConnectorsForMoved,
  setElementImageFill,
  setElementTextAnchor,
  setTableRowHeight,
  resizeTable,
  setTableCellAnchor,
  setElementParagraphFormat,
  setGroupChildParagraphFormat,
  setSlideBackground,
  setSlideBackgroundImage,
  resetSlideBackground,
  setSlideBgGraphicsHidden,
  setSlideAdvanceTime,
  setSlideHidden,
  setSlideNotes,
  setSlideTransition,
  getSections,
  setSections,
  addSection,
  renameSection,
  removeSection,
  moveSection,
  moveSlide,
  type SectionInfo,
  type ElementClipboardItem,
  type OpenedPptx,
  type Paragraph,
  type Slide,
  type TextElement,
} from '@genoffice/pptx-engine'
import { buildRenderSlide, EMU_PER_PX_96, type RenderSlide } from '@genoffice/pptx-render'
import { refineComplexWidths, shapedMetricsReady } from './shaped-metrics'
import { applyEditParagraphs, collectParagraphFormatPatches, levelsChanged } from './edit-text'
import { cfbKind, isCfbHeader } from './cfb-sniff'
import { unplayableAudioCodec } from './mp4-audio-sniff'
import type {
  AddChartOp,
  AddCommentOp,
  AddElementOp,
  AddImageBytesOp,
  AddInkOp,
  ReplacePictureBytesOp,
  AddMediaBytesOp,
  AddBlankSlideOp,
  AddSlideOp,
  PasteSlideOp,
  RepasteSlideOp,
  AddSlideWithLayoutOp,
  AddSmartArtOp,
  AddTableOp,
  ApplyThemeOp,
  HeaderFooterOp,
  SetLinkOp,
  CopyElementsOp,
  DeleteCommentOp,
  DeleteElementOp,
  EditBackgroundOp,
  EditFillOp,
  EditStrokeOp,
  FlipElementOp,
  EditPictureSrcRectOp,
  GroupElementsOp,
  UngroupElementOp,
  BatchEditTransformOp,
  EditTextOp,
  EditTransformOp,
  EditConnectorEndpointsOp,
  SetElementFontOp,
  SetElementParagraphFormatOp,
  FindReplaceOp,
  TableMergeIpcOp,
  SetSlideLayoutOp,
  SetSlideSizeOp,
  MasterEditTextOp,
  MasterEditTransformOp,
  MasterEditFillOp,
  MasterEditStrokeOp,
  MasterDeleteElementOp,
  MasterEnterResult,
  PrintSlidesOp,
  EditPictureOpacityOp,
  ExportImagesOp,
  ExportImagesResult,
  ExportPdfOp,
  ExportPdfResult,
  OpenResult,
  PasteElementsOp,
  DuplicateElementsOp,
  EditTableCellOp,
  EditTableStyleOp,
  EditChartOp,
  SetTableColWidthOp,
  SetTableRowHeightOp,
  SetTableCellAnchorOp,
  TableStructureIpcOp,
  ReorderElementOp,
  SetAdvanceTimesOp,
  SetAnimationsOp,
  SetNotesOp,
  SetSlideHiddenOp,
  SetTransitionOp,
  AddSectionOp,
  RenameSectionOp,
  RemoveSectionOp,
  MoveSectionOp,
  MoveSlideOp,
  AnimationItem,
  ShapeKey,
} from '../shared/ipc'
import { buildPrintDocumentHtml } from '../shared/print-html'

import { tm } from './i18n-main'
import { tiffToPng } from './tiff-decode'
import {
  beginHistoryBatch,
  buildAllRenderSlides,
  carryHistoryForReplacement,
  dialogParent,
  endHistoryBatch,
  getFontMetrics,
  makeMediaResolver,
  pushHistory,
  rebuildSlide,
  rebuildSlideWithReparse,
  registerAiSnapshot,
  restoreAiSnapshot,
  restoreSnapshot,
  settleStaleHistoryBatch,
  runtime,
  scheduleHistoryNotify,
  sessions,
  showChrome,
  takeSnapshot,
  windowRefs,
  type Session,
} from './session-state'
import { registerAiIpc, registerSlidesOnlyAiIpc } from './ai-ipc'
import { listPrivateFontFaces, getPrivateFontData } from './fonts'

/** One slide, copied from any deck open in this process, waiting to be pasted into another. */
let slideClipboard: { bundle: SlideBundle; png?: string } | null = null

/** The immediately preceding slide paste per webContents, so the paste-options floater can redo it with another mode. */
const lastSlidePaste = new Map<number, { afterIndex: number; undoLen: number }>()

// Cloud-generated single-page pptx: marker strings travel in pagesHtml slots; only paths issued
// by slides:cloud-page-generate are readable (the renderer can't point the reader at arbitrary files)
const CLOUD_PAGE_PREFIX = 'cloudpptx:'
const issuedCloudPages = new Set<string>()
import { registerPresenterIpc } from './presenter-show'
import { registerAttachmentIpc } from './attachments-ipc'

export {
  configureSlidesRuntime,
  setActiveSlidesWebContents,
  setSlidesShellWindow,
  setSlidesShowBleed,
} from './session-state'
export { registerAiIpc } from './ai-ipc'

/** standalone: path queued before window creation (argv/open-file) */
let pendingOpenPath: string | null = null
/** tab mode: each view queues its own path; the renderer consumes it after mounting */
const pendingByWc = new Map<number, string>()
/**
 * Renderer freeze watchdog: the freeze is sporadic and has never
 * reproduced under instrumentation, so when it does happen, capture the
 * discriminating evidence (per-process CPU/RSS, GPU feature state, and on
 * macOS native thread stacks of the renderer/GPU processes) and give
 * the user a way out. Reload restores from the main-side session, and the
 * 30s recovery draft bounds the loss for a force-quit instead.
 */
const freezeDialogOpen = new Set<number>()

async function handleRendererFreeze(wc: WebContents): Promise<void> {
  const ts = Date.now()
  try {
    const appMetrics = app.getAppMetrics()
    const diagnostics = {
      at: new Date().toISOString(),
      webContentsId: wc.id,
      sessionPath: sessions.get(wc.id)?.path ?? null,
      dirty: slidesIsDirty(wc.id),
      appMetrics,
      gpuFeatureStatus: app.getGPUFeatureStatus(),
    }
    await writeFile(
      join(app.getPath('userData'), `freeze-diagnostics-${ts}.json`),
      JSON.stringify(diagnostics, null, 2),
    )
    // macOS: native thread stacks of the renderer + GPU processes, taken from
    // outside the frozen event loop (task_for_pid based, so it works even when
    // a CDP attach suppresses the hang monitor and Runtime.evaluate is stuck).
    // This is the discriminating evidence for the freeze: a renderer main thread
    // parked in gpu::CommandBufferProxyImpl::WaitFor* confirms the GPU
    // command-buffer-wait hypothesis. Best effort — `sample` is denied for
    // hardened-runtime builds without the get-task-allow entitlement.
    if (process.platform === 'darwin') {
      const gpuPid = appMetrics.find((m) => m.type === 'GPU')?.pid
      const targets: Array<[string, number | undefined]> = [
        ['renderer', wc.getOSProcessId()],
        ['gpu', gpuPid],
      ]
      for (const [label, pid] of targets) {
        if (!pid) continue
        const out = join(app.getPath('userData'), `freeze-stacks-${ts}-${label}-${pid}.txt`)
        // Fire and forget: sampling runs for ~3s and must not delay the dialog
        execFile('/usr/bin/sample', [String(pid), '3', '-file', out], () => {})
      }
    }
  } catch {
    /* diagnostics must never make the freeze worse */
  }
  if (freezeDialogOpen.has(wc.id)) return
  freezeDialogOpen.add(wc.id)
  try {
    const parent = BrowserWindow.fromWebContents(wc)
    const options = {
      type: 'warning' as const,
      message: tm('freezeTitle'),
      detail: tm('freezeBody'),
      buttons: [tm('freezeWait'), tm('freezeReload')],
      defaultId: 0,
      cancelId: 0,
    }
    const { response } = parent
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
    if (response === 1 && !wc.isDestroyed()) {
      wc.forcefullyCrashRenderer()
      wc.reload()
    }
  } finally {
    freezeDialogOpen.delete(wc.id)
  }
}

function trackSlidesWebContents(wc: WebContents): void {
  windowRefs.activeWebContents = wc
  wc.on('unresponsive', () => void handleRendererFreeze(wc))
  // The AI panel opens links via window.open; route them to the system
  // browser instead of spawning an in-app window with remote content.
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url)
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    // Untitled recovery draft: cleaned up on a clean close, kept when the session died dirty
    const s = sessions.get(wc.id)
    if (s && !sessionDirty(s)) dropUntitledRecovery(wc.id)
    else untitledRecovery.delete(wc.id)
    sessions.delete(wc.id)
    pendingByWc.delete(wc.id)
    clipboards.delete(wc.id)
    lastSlidePaste.delete(wc.id)
    closeSaveWaiters.get(wc.id)?.(false)
    closeSaveWaiters.delete(wc.id)
    autoSavePrefByWc.delete(wc.id)
    if (windowRefs.activeWebContents === wc) windowRefs.activeWebContents = null
  })
}

// ── In-app element clipboard (isolated per webContents; pasteCount drives cascading offset) ─
const clipboards = new Map<number, { items: ElementClipboardItem[]; pasteCount: number }>()

/** Shell hook: a view opened a file (including ⌘O inside a tab) — used to update tab titles and de-duplicate paths */
let slidesOpenedHook: ((wc: WebContents, path: string) => void) | null = null
export function setSlidesOpenedHook(fn: ((wc: WebContents, path: string) => void) | null): void {
  slidesOpenedHook = fn
}

const RECENT_PATH = () => join(app.getPath('userData'), 'slides-recent.json')

/** Comment author name: system username, falling back to a generic "User" label. */
function commentAuthorName(): string {
  try {
    return userInfo().username || 'User'
  } catch {
    return 'User'
  }
}

async function readRecent(): Promise<string[]> {
  try {
    const raw = await readFile(RECENT_PATH(), 'utf8')
    return (JSON.parse(raw) as string[]).filter((p) => existsSync(p))
  } catch {
    return []
  }
}

async function pushRecent(path: string): Promise<void> {
  const cur = await readRecent()
  const next = [path, ...cur.filter((p) => p !== path)].slice(0, 10)
  try {
    await writeFile(RECENT_PATH(), JSON.stringify(next), 'utf8')
  } catch {
    /* best-effort */
  }
}

/** File renamed externally (shell Home list rename): swap the old path in the recent list for the new one (keeping its position). */
export async function replaceSlidesRecentFile(oldPath: string, newPath: string): Promise<void> {
  try {
    // Do not use readRecent(): it filters out old paths that no longer exist, so the map would miss
    const raw = await readFile(RECENT_PATH(), 'utf8')
    const cur = JSON.parse(raw) as string[]
    await writeFile(
      RECENT_PATH(),
      JSON.stringify(cur.map((p) => (p === oldPath ? newPath : p))),
      'utf8',
    )
  } catch {
    /* best-effort */
  }
}

/** Shell notification: an open view's file was renamed — sync the session path (subsequent
 *  saves write the new file) and push to the renderer to update the editor title bar. */
export function slidesFileRenamed(wc: WebContents, oldPath: string, newPath: string): void {
  const session = sessions.get(wc.id)
  if (session && session.path === oldPath) session.path = newPath
  wc.send('slides:renamed', newPath)
}

// ── Autosave (crash recovery): dirty sessions write a recovery copy every 30s; a normal save cleans it up ──
const autosaveDir = () => join(app.getPath('userData'), 'slides-autosave')
const autosavePathFor = (filePath: string) =>
  join(autosaveDir(), `${createHash('sha1').update(filePath).digest('hex').slice(0, 16)}.pptx`)

function sessionDirty(session: Session): boolean {
  return (
    !!session.metaDirty ||
    session.opened.deck.slides.some(
      (s) => s.structureDirty || s.elements.some((el) => el.dirty || el.dirtyTransform),
    )
  )
}

/**
 * Ticks to skip after a failed recovery copy, per deck. Retrying every 30s just
 * repeats an expensive failure, but disabling the safety net for the rest of the
 * session was worse: on a heavy deck one slow serialization used to remove crash
 * recovery permanently and silently. Back off instead, and keep the
 * skip count so a deck that always fails only pays for it every ~5 minutes.
 */
const autosaveBackoff = new Map<string, number>()
const AUTOSAVE_BACKOFF_TICKS = 10
let autosaveRunning = false

/**
 * Recovery drafts for never-saved decks (wcId → visible path in <Documents>/GenOffice):
 * the sha1-keyed recovery copy needs session.path, so before the first save a freeze or
 * crash used to lose everything. Removed on save, explicit discard, or clean close.
 */
const untitledRecovery = new Map<number, string>()

function dropUntitledRecovery(wcId: number): void {
  const draft = untitledRecovery.get(wcId)
  if (draft) void rm(draft, { force: true }).catch(() => {})
  untitledRecovery.delete(wcId)
}

setInterval(() => {
  if (autosaveRunning) return
  autosaveRunning = true
  void (async () => {
    for (const [wcId, session] of sessions.entries()) {
      if (session.masterEdit || !sessionDirty(session)) continue
      let target: string
      if (session.path) {
        target = autosavePathFor(session.path)
      } else {
        let draft = untitledRecovery.get(wcId)
        if (!draft) {
          draft = join(getDraftsDir(), newDraftFilename())
          untitledRecovery.set(wcId, draft)
        }
        target = draft
      }
      const backoffKey = session.path ?? target
      const skip = autosaveBackoff.get(backoffKey) ?? 0
      if (skip > 0) {
        autosaveBackoff.set(backoffKey, skip - 1)
        continue
      }
      try {
        await mkdir(dirname(target), { recursive: true })
        await savePptxToFile(session.opened, target)
        autosaveBackoff.delete(backoffKey)
      } catch (error) {
        autosaveBackoff.set(backoffKey, AUTOSAVE_BACKOFF_TICKS)
        console.warn('[slides] autosave failed, retrying in ~5 min:', error)
      }
    }
  })().finally(() => {
    autosaveRunning = false
  })
}, 30_000)

// ── Close guard (aligned with sheets/pdf): dirty sessions prompt Save/Don't Save/Cancel before closing a tab/window ──
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
/** Autosave toggle mirrored from the renderer: files with it on save silently on close and proceed, no dialog */
const autoSavePrefByWc = new Map<number, boolean>()

ipcMain.on('slides:autosave-pref', (event, on: unknown) => {
  autoSavePrefByWc.set(event.sender.id, on === true)
})

ipcMain.on('slides:close-save-result', (event, ok: unknown) => {
  const waiter = closeSaveWaiters.get(event.sender.id)
  if (!waiter) return
  closeSaveWaiters.delete(event.sender.id)
  waiter(ok === true)
})

/** Ask the renderer to run the full save flow and await the result (failure/timeout = false). */
function requestRendererSave(contents: WebContents): Promise<boolean> {
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send('slides:close-save-request')
  })
}

export function slidesIsDirty(webContentsId: number): boolean {
  const session = sessions.get(webContentsId)
  return !!session && sessionDirty(session)
}

/**
 * Close guard for the slides renderer: true means proceed with closing.
 * Clean -> true; with changes -> Save/Don't Save/Cancel. Choosing Save asks the renderer to run
 * the existing save flow (flushNotes + adoptSavedSlides + Save As dialog for untitled) and
 * awaits the result; on failure/timeout stay open.
 */
export async function requestSlidesClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  if (!slidesIsDirty(contents.id) || contents.isDestroyed()) return true
  // Autosave on and a path exists: save silently and proceed without bothering the user; only a failed save falls through to the dialog
  if (autoSavePrefByWc.get(contents.id) && sessions.get(contents.id)?.path) {
    if (await requestRendererSave(contents)) return true
  }
  const options = {
    type: 'warning' as const,
    message: tm('closeUnsavedMsg'),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('menuSave'), tm('btnDontSave'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const { response } =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (response === 2) return false
  if (response === 1) {
    // User explicitly discarded changes: also delete the recovery copy, so the next open does not show a pointless recovery prompt
    const session = sessions.get(contents.id)
    if (session?.path) void rm(autosavePathFor(session.path), { force: true }).catch(() => {})
    dropUntitledRecovery(contents.id)
    return true
  }
  return requestRendererSave(contents)
}

/** On open, if a recovery copy newer than the original exists, ask whether to restore (still points at the original path; only save persists it). */
async function maybeRecoverBytes(
  path: string,
  original: Uint8Array,
): Promise<{ bytes: Uint8Array; recovered: boolean }> {
  const asPath = autosavePathFor(path)
  try {
    const [asStat, origStat] = await Promise.all([stat(asPath), stat(path)])
    if (asStat.mtimeMs <= origStat.mtimeMs) {
      await rm(asPath, { force: true })
      return { bytes: original, recovered: false }
    }
  } catch {
    return { bytes: original, recovered: false }
  }
  const parent = dialogParent()
  const options = {
    type: 'question' as const,
    buttons: [tm('autosaveRestore'), tm('autosaveDiscard')],
    defaultId: 0,
    cancelId: 1,
    message: tm('autosaveFoundTitle'),
    detail: tm('autosaveFoundBody'),
  }
  const r = parent
    ? await dialog.showMessageBox(parent, options)
    : await dialog.showMessageBox(options)
  if (r.response === 0) {
    const bytes = await readFile(asPath)
    return { bytes: new Uint8Array(bytes), recovered: true }
  }
  await rm(asPath, { force: true })
  return { bytes: original, recovered: false }
}

/**
 * .ppt (97-2003 binary compound document) and encrypted OOXML are unsupported: show an actionable message instead of a parse error.
 * Detection uses the magic number rather than the extension -- a binary ppt with a renamed suffix is caught too. A CFB containing an
 * EncryptedPackage stream is a password-protected pptx and gets dedicated copy (instead of being mislabeled as the legacy format).
 */
async function rejectLegacyPpt(path: string): Promise<boolean> {
  let head: Buffer
  try {
    const fh = await open(path, 'r')
    try {
      head = Buffer.alloc(8)
      await fh.read(head, 0, 8, 0)
    } finally {
      await fh.close()
    }
  } catch {
    return false
  }
  if (!isCfbHeader(head)) return false
  let kind: 'legacy' | 'encrypted' = 'legacy'
  try {
    kind = cfbKind(await readFile(path)) ?? 'legacy'
  } catch {
    // on read failure, fall back to the legacy-format message
  }
  const parent = dialogParent()
  const options = {
    type: 'warning' as const,
    buttons: [tm('legacyPptOk')],
    message: tm(kind === 'encrypted' ? 'encryptedPptxTitle' : 'legacyPptTitle'),
    detail: tm(kind === 'encrypted' ? 'encryptedPptxBody' : 'legacyPptBody'),
  }
  if (parent) await dialog.showMessageBox(parent, options)
  else await dialog.showMessageBox(options)
  return true
}

async function openAndBuild(
  wc: WebContents,
  path: string,
  fitWidthPx: number,
): Promise<OpenResult> {
  const raw = await readFile(path)
  const { bytes, recovered } = await maybeRecoverBytes(path, new Uint8Array(raw))
  await shapedMetricsReady() // Lay out only after complex-script shaped metrics are ready, avoiding an init race falling back to estimation
  const opened = await openPptx(bytes)
  sessions.set(wc.id, {
    path,
    opened,
    fitWidthPx,
    undoStack: [],
    redoStack: [],
    ...(recovered ? { metaDirty: true } : {}),
  })
  scheduleHistoryNotify(sessions.get(wc.id)!)
  await pushRecent(path)
  slidesOpenedHook?.(wc, path)
  let slides = buildAllRenderSlides(opened, fitWidthPx)
  // If the first layout pass had complex-script misses (Arabic/Thai etc.), re-lay out once with renderer-measured widths
  if (await refineComplexWidths(wc)) slides = buildAllRenderSlides(opened, fitWidthPx)
  return {
    path,
    slides,
    size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy },
    defaultFont: deckDefaultFont(opened),
  }
}

/** Directory where AI-generated drafts are saved: the configurable default save folder (falls back to <Documents>/GenOffice) */
function getDraftsDir(): string {
  return configuredDefaultSaveDir(app)
}

/** Fallback draft filename: <untitled label>-YYYYMMDD-HHmmss.pptx */
function newDraftFilename(): string {
  const d = new Date()
  const pad = (n: number, w = 2) => String(n).padStart(w, '0')
  const date = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`
  const time = `${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`
  return `${tm('untitledDraft')}-${date}-${time}.pptx`
}

/** Sanitize an AI-provided topic/title into a safe filename base: strip illegal path chars, collapse whitespace, cap length; null if invalid. */
function sanitizeDraftBaseName(raw: string | undefined): string | null {
  if (!raw) return null
  const cleaned = raw
    // eslint-disable-next-line no-control-regex -- stripping control chars is the point here
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    // Strip leading/trailing dots (Windows disallows a trailing dot; a hidden-file prefix is meaningless here)
    .replace(/^\.+|\.+$/g, '')
    .trim()
  if (!cleaned) return null
  return cleaned.length > 40 ? cleaned.slice(0, 40).trim() : cleaned
}

/** Pick a draft path from deckName: append -2/-3… if a same-named file exists; fall back to timestamp naming without a valid deckName. */
function pickDraftPath(draftsDir: string, deckName?: string): string {
  const base = sanitizeDraftBaseName(deckName)
  if (base) {
    let candidate = join(draftsDir, `${base}.pptx`)
    for (let i = 2; existsSync(candidate) && i < 100; i++) {
      candidate = join(draftsDir, `${base}-${i}.pptx`)
    }
    if (!existsSync(candidate)) return candidate
  }
  return join(draftsDir, newDraftFilename())
}

/**
 * Auto-save the draft to <Documents>/GenOffice/<name>.pptx after AI generation completes.
 * Append mode reuses the session's existing draft path (overwrite); replace mode generates a
 * new filename. On successful write, update session.path, pushRecent, slidesOpenedHook.
 * On write failure, degrade silently (console.warn) without blocking the in-memory session.
 */
async function saveDraftAfterGenerate(
  wc: WebContents,
  session: Session,
  bytes: Uint8Array,
  mode: 'replace' | 'append',
  deckName?: string,
): Promise<void> {
  try {
    const draftsDir = getDraftsDir()
    // Ensure the directory exists
    if (!existsSync(draftsDir)) mkdirSync(draftsDir, { recursive: true })

    // Append mode: overwrite if the session already has a draft path; otherwise create a new file too
    let draftPath: string
    if (mode === 'append' && session.path && session.path.startsWith(draftsDir)) {
      draftPath = session.path
    } else {
      draftPath = pickDraftPath(draftsDir, deckName)
    }

    await writeFile(draftPath, Buffer.from(bytes))
    session.path = draftPath
    await pushRecent(draftPath)
    slidesOpenedHook?.(wc, draftPath)
  } catch (err) {
    console.warn(
      '[slides] Failed to persist AI-generated draft to disk; the in-memory session still works:',
      err,
    )
  }
}

/** Theme body (minor) Latin font: fallback shown in the ribbon font box when the selection has no text element. */
function deckDefaultFont(opened: OpenedPptx): string | undefined {
  try {
    const slidePath = opened.archive.readPresentation().slidePaths[0]
    if (!slidePath) return undefined
    const themePath = opened.archive.resolveSlideChain(slidePath).themePath
    const xml = themePath ? opened.archive.readText(themePath) : undefined
    return xml ? parseTheme(xml).minorFont : undefined
  } catch {
    return undefined
  }
}

function findEl(slide: Slide, sourceId: string): TextElement | undefined {
  const el = slide.elements.find((e) => e.id === sourceId)
  if (el && (el.type === 'text' || el.type === 'shape')) return el as TextElement
  return undefined
}

/**
 * spAutoFit (autofit='resize', "resize shape to fit text"): after a text change, the box height
 * grows/shrinks with the content and is written back to cy. rendered = the
 * rebuilt result after this change; when the height changed, update the transform and rebuild
 * once more. Top-level elements only (group children use a different coordinate system, skip).
 */
function applyAutofitResize(
  session: Session,
  slideIndex: number,
  sourceId: string,
  rendered: RenderSlide | null,
): RenderSlide | null {
  if (!rendered) return rendered
  const slide = session.opened.deck.slides[slideIndex]
  const el = slide ? findEl(slide, sourceId) : undefined
  if (!el?.text || el.text.autofit !== 'resize') return rendered
  const node = rendered.nodes.find((n) => n.sourceId === sourceId)
  if (!node || (node.type !== 'shape' && node.type !== 'text') || !node.text) return rendered
  const needH = node.text.contentHeight + node.text.insets.t + node.text.insets.b
  if (Math.abs(needH - node.box.h) < 1) return rendered
  const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
  const scale = session.fitWidthPx / baseWidthPx
  el.transform = {
    ...el.transform,
    offset: {
      ...el.transform.offset,
      cy: Math.max(Math.round((needH / scale) * EMU_PER_PX_96), 1),
    },
  }
  el.dirtyTransform = true
  return rebuildSlide(session, slideIndex)
}

/**
 * normAutofit fontScale write-back: when the shrink ratio the layout actually used after a text
 * edit (≤ the stored cap, shrink-only) differs from the stored model value, sync the model and
 * patch the bodyPr attribute — only then does PowerPoint show the same size on open.
 * Triggered only by text edits (resize gestures do not write: the layout cap locks the stored
 * value, and writing back during a gesture would ratchet one way); top-level elements only.
 */
function syncAutofitScale(
  session: Session,
  slideIndex: number,
  sourceId: string,
  rendered: RenderSlide | null,
): RenderSlide | null {
  if (!rendered) return rendered
  const slide = session.opened.deck.slides[slideIndex]
  const el = slide ? findEl(slide, sourceId) : undefined
  if (!el?.text || el.text.autofit !== 'shrink') return rendered
  const node = rendered.nodes.find((n) => n.sourceId === sourceId)
  if (!node || (node.type !== 'shape' && node.type !== 'text') || !node.text) return rendered
  const effective = node.text.fontScale
  const effectiveRed = node.text.lnSpcReduction ?? 0
  if (
    Math.abs(effective - (el.text.fontScale ?? 1)) < 0.005 &&
    Math.abs(effectiveRed - (el.text.lnSpcReduction ?? 0)) < 0.005
  )
    return rendered
  el.text.fontScale = effective
  if (effectiveRed) el.text.lnSpcReduction = effectiveRed
  else delete el.text.lnSpcReduction
  el.anchor.originalXml = patchBodyPrAutofit(el.anchor.originalXml, effective, effectiveRed)
  slide!.structureDirty = true
  return rendered // The layout already rendered with the effective value; no rebuild needed
}

/** Legacy fixed color schemes (AI tools/old files still pass these keys; kept as fallback). */
const CHART_COLOR_SCHEMES: Record<string, string[]> = {
  default: [],
  blue: ['#2E75B6', '#4472C4', '#5B9BD5', '#70AD47', '#ED7D31'],
  warm: ['#ED7D31', '#FFC000', '#FF0000', '#C55A11', '#833C00'],
  cool: ['#0070C0', '#00B0F0', '#00B0A0', '#7030A0', '#2E75B6'],
  mono: ['#404040', '#666666', '#888888', '#AAAAAA', '#CCCCCC'],
}

/** PowerPoint default theme accent sequence (fallback when the deck has no theme colors). */
const FALLBACK_ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

/** Mix a hex color with black/white by ratio (for mono-gradient steps). */
function mixHex(hex: string, target: number, ratio: number): string {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(hex)
  if (!m) return hex
  const v = parseInt(m[1]!, 16)
  const ch = (x: number) => Math.round(x + (target - x) * ratio)
  const r = ch((v >> 16) & 255)
  const g = ch((v >> 8) & 255)
  const b = ch(v & 255)
  return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, '0').toUpperCase()}`
}

/** Current deck's theme accent1..6 (read from the theme part of the first slide's inheritance chain). */
function deckAccents(opened: OpenedPptx): string[] {
  const slide = opened.deck.slides[0]
  if (!slide) return FALLBACK_ACCENTS
  try {
    const chain = opened.archive.resolveSlideChain(slide.path)
    const xml = chain.themePath ? opened.archive.readText(chain.themePath) : null
    const colors = xml ? parseTheme(xml).colors : undefined
    const acc = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
      .map((k) => colors?.[k])
      .filter((c): c is string => !!c)
    return acc.length >= 3 ? acc : FALLBACK_ACCENTS
  } catch {
    return FALLBACK_ACCENTS
  }
}

/** Theme-derived color schemes for the chart "Change Colors" gallery: two colorful sets + one mono gradient per accent. */
function chartColorSchemes(
  opened: OpenedPptx,
): Array<{ key: string; label: string; colors: string[] }> {
  const acc = deckAccents(opened)
  const rot = [...acc.slice(3), ...acc.slice(0, 3)]
  const mono = (c: string) => [
    mixHex(c, 0, 0.25),
    c,
    mixHex(c, 255, 0.25),
    mixHex(c, 255, 0.45),
    mixHex(c, 255, 0.65),
  ]
  return [
    { key: 'default', label: tm('schemeThemeDefault'), colors: [] },
    { key: 'colorful', label: tm('schemeColorful'), colors: acc },
    { key: 'colorful2', label: tm('schemeColorful2'), colors: rot },
    ...acc.map((c, i) => ({
      key: `mono-accent${i + 1}`,
      label: tm('schemeMono', { n: i + 1 }),
      colors: mono(c),
    })),
  ]
}

let ipcRegistered = false

export function registerSlidesIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  // shared with the other editor modules — last (identical) registration wins
  ipcMain.removeHandler('app:get-language')
  ipcMain.handle('app:get-language', () => getUiLang())

  // Screen recording: source dispatch for the renderer's navigator.mediaDevices.getDisplayMedia.
  // macOS prefers the system picker (with its permission flow), falling back to the first screen.
  void app.whenReady().then(() => {
    try {
      electronSession.defaultSession.setDisplayMediaRequestHandler(
        (_request, callback) => {
          desktopCapturer
            .getSources({ types: ['screen', 'window'] })
            .then((sources) => {
              if (sources[0]) callback({ video: sources[0] })
              else callback({})
            })
            .catch(() => callback({}))
        },
        { useSystemPicker: true },
      )
    } catch {
      /* Older Electron lacks this API: the screen-record button will get no stream and report failure */
    }
  })

  ipcMain.handle('slides:private-font-faces', () => listPrivateFontFaces())
  ipcMain.handle('slides:private-font-data', (_e, id: string) => getPrivateFontData(id))

  ipcMain.handle('slides:open', async (e, fitWidthPx: number) => {
    const parent = dialogParent()
    const options = {
      properties: ['openFile' as const],
      filters: [{ name: 'PowerPoint', extensions: ['pptx', 'ppt'] }],
    }
    const r = await showOpenDialogWithMemory(dialog, parent, options)
    if (r.canceled || !r.filePaths[0]) return null
    if (await rejectLegacyPpt(r.filePaths[0])) return null
    return openAndBuild(e.sender, r.filePaths[0], fitWidthPx)
  })

  ipcMain.handle('slides:open-path', async (e, path: string, fitWidthPx: number) => {
    if (!path || !existsSync(path)) return null
    if (await rejectLegacyPpt(path)) return null
    return openAndBuild(e.sender, path, fitWidthPx)
  })

  ipcMain.handle('slides:consume-pending-open', async (e, fitWidthPx: number) => {
    // renderer app just mounted: safe to reveal the vibrancy material behind
    // the (now painted) page without flashing raw desktop during load
    vibFlip.get(e.sender.id)?.('#00000000')
    const queued = pendingByWc.get(e.sender.id) ?? pendingOpenPath
    if (queued && existsSync(queued)) {
      // Clear the queue only after a successful open: keep it on parse failure or a mid-flight renderer reload, so a remount can retry
      const result = await openAndBuild(e.sender, queued, fitWidthPx)
      if (pendingByWc.get(e.sender.id) === queued) pendingByWc.delete(e.sender.id)
      if (pendingOpenPath === queued) pendingOpenPath = null
      return result
    }
    // No queued path but the main process already has a session (remount after an HMR full
    // reload/crash recovery) -> restore from the session; otherwise the document is lost leaving
    // only the start screen, and reopening the same file just activates this empty tab with no
    // way to self-heal
    const session = sessions.get(e.sender.id)
    if (session) {
      session.fitWidthPx = fitWidthPx
      return {
        path: session.path,
        slides: buildAllRenderSlides(session.opened, fitWidthPx),
        size: { cx: session.opened.deck.size.cx, cy: session.opened.deck.size.cy },
        defaultFont: deckDefaultFont(session.opened),
      } satisfies OpenResult
    }
    return null
  })

  ipcMain.handle('slides:edit-text', (e, op: EditTextOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    // In-group editing: after updating a child's paragraphs, patch the text of the child slice inside the group's originalXml
    if (op.groupId) {
      const found = findGroupChild(slide, op.groupId, op.sourceId)
      const child = found?.child
      if (!child || (child.type !== 'text' && child.type !== 'shape')) return null
      const textChild = child as TextElement
      if (!textChild.text) return null
      pushHistory(session)
      textChild.text.paragraphs = applyEditParagraphs(textChild.text.paragraphs, op.paragraphs)
      ensureRunLinkRels(session.opened, op.slideIndex, textChild.text.paragraphs)
      if (!patchGroupChildText(slide, op.groupId, textChild)) {
        restoreSnapshot(session, session.undoStack.pop()!) // Slice not located: roll back the already-modified model
        return null
      }
      for (const { index, patch } of collectParagraphFormatPatches(op.paragraphs)) {
        setGroupChildParagraphFormat(slide, op.groupId, op.sourceId, patch, [index])
      }
      return rebuildSlide(session, op.slideIndex)
    }
    const el = findEl(slide, op.sourceId)
    if (!el || !el.text) return null
    pushHistory(session)
    // Run-level rich-text rebuild: srcPara/srcRun back-tracing + preserving unedited fields, see applyEditParagraphs
    const levelDirty = levelsChanged(el.text.paragraphs, op.paragraphs)
    el.text.paragraphs = applyEditParagraphs(el.text.paragraphs, op.paragraphs)
    ensureRunLinkRels(session.opened, op.slideIndex, el.text.paragraphs)
    el.dirty = true
    // Per-paragraph bullets/spacing marked on the editor selection
    for (const { index, patch } of collectParagraphFormatPatches(op.paragraphs)) {
      setElementParagraphFormat(slide, op.sourceId, patch, [index])
    }
    if (levelDirty) {
      // Level changes affect inheritance (font size/bullet/indent take master defaults by lvl); bake into bytes then reparse
      el.dirtyPPr = { ...el.dirtyPPr, level: true, indents: true }
      materializeSlide(session.opened, op.slideIndex)
      return rebuildSlide(session, op.slideIndex)
    }
    const rendered = applyAutofitResize(
      session,
      op.slideIndex,
      op.sourceId,
      rebuildSlide(session, op.slideIndex),
    )
    return syncAutofitScale(session, op.slideIndex, op.sourceId, rendered)
  })

  ipcMain.handle('slides:set-element-font', (e, op: SetElementFontOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    let changed = false
    for (const id of op.sourceIds) {
      const ok = op.groupId
        ? setGroupChildFont(slide, op.groupId, id, {
            fontFamily: op.fontFamily,
            fontSizePt: op.fontSizePt,
            strike: op.strike,
            bold: op.bold,
            italic: op.italic,
            underline: op.underline,
            color: op.color,
          })
        : setElementFont(slide, id, {
            fontFamily: op.fontFamily,
            fontSizePt: op.fontSizePt,
            strike: op.strike,
            bold: op.bold,
            italic: op.italic,
            underline: op.underline,
            color: op.color,
          })
      if (ok) changed = true
    }
    if (!changed) {
      session.undoStack.pop() // All non-text elements (images etc.): nothing happened, pop the just-pushed history
      return null
    }
    let rendered = rebuildSlide(session, op.slideIndex)
    for (const id of op.sourceIds) {
      rendered = applyAutofitResize(session, op.slideIndex, id, rendered)
      rendered = syncAutofitScale(session, op.slideIndex, id, rendered)
    }
    return rendered
  })

  ipcMain.handle('slides:set-element-paragraph-format', (e, op: SetElementParagraphFormatOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    const patch = {
      bullet: op.bullet,
      bulletChar: op.bulletChar,
      bulletHangEmu: op.bulletHangEmu,
      bulletSizePct: op.bulletSizePct,
      bulletColor: op.bulletColor,
      lineSpacingPct: op.lineSpacingPct,
      spaceBeforePt: op.spaceBeforePt,
      spaceAfterPt: op.spaceAfterPt,
      align: op.align,
      indentDelta: op.indentDelta,
    }
    let changed = false
    for (const id of op.sourceIds) {
      const ok = op.groupId
        ? setGroupChildParagraphFormat(slide, op.groupId, id, patch)
        : setElementParagraphFormat(slide, id, patch)
      if (ok) changed = true
    }
    if (!changed) {
      session.undoStack.pop()
      return null
    }
    if (op.indentDelta) {
      // Level changes affect inherited defaults; bake into bytes then reparse
      materializeSlide(session.opened, op.slideIndex)
      return rebuildSlide(session, op.slideIndex)
    }
    let rendered = rebuildSlide(session, op.slideIndex)
    for (const id of op.sourceIds) {
      rendered = applyAutofitResize(session, op.slideIndex, id, rendered)
      rendered = syncAutofitScale(session, op.slideIndex, id, rendered)
    }
    return rendered
  })

  ipcMain.handle('slides:edit-transform', (e, op: EditTransformOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const el = op.groupId ? null : slide.elements.find((x) => x.id === op.sourceId)
    const grpChild = op.groupId ? findGroupChild(slide, op.groupId, op.sourceId) : null
    if (!el && !grpChild) return null
    // Undo semantics for preview gestures: one whole drag = one undo step.
    // The first preview pushes a pre-gesture snapshot; later previews and the final commit do not.
    if (op.preview) {
      if (!session.transformPreview) {
        pushHistory(session)
        session.transformPreview = true
      }
    } else if (session.transformPreview) {
      session.transformPreview = false
    } else {
      pushHistory(session)
    }
    // px -> EMU (inverting the viewport scale)
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    // In-group editing: the pixel box is in group-local coords (with ext/chExt scaling baked in); divide out the group scale first, then convert back to the child EMU coordinate system
    if (grpChild) {
      const ch = grpChild.grp.childOffset
      const chX = ch?.x ?? grpChild.grp.transform.offset.x
      const chY = ch?.y ?? grpChild.grp.transform.offset.y
      const gExt = grpChild.grp.transform.offset
      const gsx = ch?.cx ? gExt.cx / ch.cx : 1
      const gsy = ch?.cy ? gExt.cy / ch.cy : 1
      const ok = editGroupChildTransform(
        slide,
        op.groupId!,
        op.sourceId,
        {
          x: toEmu(op.xPx / gsx) + chX,
          y: toEmu(op.yPx / gsy) + chY,
          cx: toEmu(op.wPx / gsx),
          cy: toEmu(op.hPx / gsy),
        },
        op.rotationDeg,
      )
      if (!ok) {
        session.undoStack.pop() // Slice not located: model untouched, pop the just-pushed history
        return null
      }
      return rebuildSlide(session, op.slideIndex)
    }
    // Tables: redistribute gridCol widths / tr heights so the file matches the
    // frame instead of keeping the old grid under a new a:ext
    const isTable = el!.type === 'table'
    if (isTable) resizeTable(slide, op.sourceId, toEmu(op.wPx), toEmu(op.hPx))
    el!.transform = {
      ...el!.transform,
      offset: {
        x: toEmu(op.xPx),
        y: toEmu(op.yPx),
        // resizeTable synced cx/cy to the redistributed sums
        cx: isTable ? el!.transform.offset.cx : toEmu(op.wPx),
        cy: isTable ? el!.transform.offset.cy : toEmu(op.hPx),
      },
      rot: Math.round(op.rotationDeg * 60000),
    }
    el!.dirtyTransform = true
    updateConnectorsForMoved(slide, [op.sourceId])
    return rebuildSlide(session, op.slideIndex)
  })

  // Connector endpoint drag: box+flip re-derived from the two endpoints;
  // attach/detach writes a:stCxn/a:endCxn so the connector follows later shape moves
  ipcMain.handle('slides:edit-connector-endpoints', (e, op: EditConnectorEndpointsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const el = slide.elements.find((x) => x.id === op.sourceId)
    if (!el) return null
    pushHistory(session)
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const p1 = { x: toEmu(op.x1Px), y: toEmu(op.y1Px) }
    const p2 = { x: toEmu(op.x2Px), y: toEmu(op.y2Px) }
    el.transform = {
      ...el.transform,
      offset: {
        x: Math.min(p1.x, p2.x),
        y: Math.min(p1.y, p2.y),
        cx: Math.abs(p2.x - p1.x),
        cy: Math.abs(p2.y - p1.y),
      },
      rot: 0,
      flipH: p1.x > p2.x,
      flipV: p1.y > p2.y,
    }
    el.dirtyTransform = true
    const toRef = (
      v: { targetId: string; idx: number } | null | undefined,
    ): { id: number; idx: number } | null | undefined => {
      if (v === undefined) return undefined
      if (v === null) return null
      const target = slide.elements.find((x) => x.id === v.targetId)
      const spid = target ? elementSpid(target) : null
      return spid != null ? { id: spid, idx: v.idx } : null
    }
    setElementConnection(slide, op.sourceId, { start: toRef(op.start), end: toRef(op.end) })
    return rebuildSlide(session, op.slideIndex)
  })

  // Read-only: RenderSlide for every page of the current session (E2E driver/debug use, no state change)
  ipcMain.handle('slides:get-render-slides', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    return session.opened.deck.slides.map((_, i) => rebuildSlide(session, i))
  })

  ipcMain.handle('slides:batch-edit-transform', (e, op: BatchEditTransformOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    // Validate: every element must exist
    const pairs: Array<{ el: (typeof slide.elements)[0]; item: BatchEditTransformOp['items'][0] }> =
      []
    for (const item of op.items) {
      const el = slide.elements.find((x) => x.id === item.sourceId)
      if (!el) return null
      pairs.push({ el, item })
    }
    pushHistory(session)
    for (const { el, item } of pairs) {
      el.transform = {
        ...el.transform,
        offset: {
          x: toEmu(item.xPx),
          y: toEmu(item.yPx),
          cx: toEmu(item.wPx),
          cy: toEmu(item.hPx),
        },
        rot: Math.round(item.rotationDeg * 60000),
      }
      el.dirtyTransform = true
    }
    updateConnectorsForMoved(
      slide,
      op.items.map((i) => i.sourceId),
    )
    return rebuildSlide(session, op.slideIndex)
  })
  // ── Cloud single-page generation (gsk slide_generate): brief → cloud HTML+conversion → one-slide
  // pptx saved to a temp file. Returns a marker string that flows through the same pagesHtml slots
  // as locally generated HTML; slides:html-to-pptx recognizes it and reads the bytes instead of
  // converting. Enabled when gsk is logged in; GENOFFICE_CLOUD_SLIDE=0 is the kill switch.
  const cloudSlideEnabled = () => process.env.GENOFFICE_CLOUD_SLIDE !== '0' && !!gskApiKey()

  ipcMain.handle('slides:cloud-gen-status', () => ({ enabled: cloudSlideEnabled() }))

  ipcMain.handle(
    'slides:cloud-page-generate',
    async (
      _e,
      op: {
        brief: string
        title?: string
        styleSkill?: string
        deckContext?: Record<string, unknown>
        images?: { url: string; caption?: string }[]
        width?: number
        height?: number
      },
    ): Promise<{ ok: boolean; marker?: string; error?: string }> => {
      if (!cloudSlideEnabled()) return { ok: false, error: 'cloud slide generation is disabled' }
      try {
        // ultra = opus-class model, matching the local path's quality tier; GENOFFICE_CLOUD_SLIDE_TIER=standard opts down
        const tier = process.env.GENOFFICE_CLOUD_SLIDE_TIER === 'standard' ? 'standard' : 'ultra'
        const started = Date.now()
        const { bytes, model } = await gskSlideGenerate({
          tier,
          brief: String(op.brief ?? ''),
          title: op.title ? String(op.title) : undefined,
          styleSkill: op.styleSkill ? String(op.styleSkill) : undefined,
          deckContext: op.deckContext,
          images: Array.isArray(op.images) ? op.images : undefined,
          width: op.width,
          height: op.height,
        })
        console.log(
          `[cloud-slide] page generated: tier=${tier} model=${model} bytes=${bytes.length} ms=${Date.now() - started}`,
        )
        const dir = join(app.getPath('temp'), 'genoffice-cloud-pages')
        mkdirSync(dir, { recursive: true })
        const path = join(dir, `${randomUUID()}.pptx`)
        await writeFile(path, bytes)
        issuedCloudPages.add(path)
        return { ok: true, marker: CLOUD_PAGE_PREFIX + path }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    'slides:html-to-pptx',
    async (
      e,
      pagesHtml: string[],
      fitWidthPx: number,
      mode?: 'replace' | 'append' | 'replace_at' | 'insert_at',
      atIndex?: number,
      deckName?: string,
    ): Promise<
      | (OpenResult & {
          appendedFrom?: number
          replacedIndex?: number
          insertedIndex?: number
          fallbackReason?: string
          imageFailures?: { page: number; url: string }[]
        })
      | { error: string }
    > => {
      // Every page arrives as a cloud marker (cloudpptx:<path> written by
      // slides:cloud-page-generate, pointing at a one-slide pptx temp file); this handler only
      // reads and lands the bytes.
      // replace: assemble the whole batch into one multi-page pptx as the new deck base.
      // append: merge the "new pages" one by one into the existing deck via mergeSlideFromPptx
      // (earlier pages are untouched).
      const readCloudPage = async (marker: string): Promise<{ bytes: Uint8Array }> => {
        if (!marker.startsWith(CLOUD_PAGE_PREFIX)) throw new Error('expected a cloud page marker')
        const path = marker.slice(CLOUD_PAGE_PREFIX.length)
        if (!issuedCloudPages.has(path)) throw new Error('unknown cloud page marker')
        return { bytes: new Uint8Array(await readFile(path)) }
      }
      const assembleDeck = async (): Promise<{ bytes: Uint8Array }> => {
        const perPage = await Promise.all(pagesHtml.map(readCloudPage))
        const base = await openPptx(perPage[0]!.bytes)
        for (const one of perPage.slice(1)) await mergeSlideFromPptx(base, one.bytes)
        for (const s of base.deck.slides) promoteSlideBackground(s, base.deck.size)
        return { bytes: await savePptx(base) }
      }

      try {
        // Append: convert only the "new pages" and merge them one by one into the existing
        // in-memory deck via mergeSlideFromPptx. Already-landed pages stay untouched
        // (O(N) rather than O(N²)); no dependency on stored PageVisualData.
        if (mode === 'append') {
          const existing = sessions.get(e.sender.id)
          if (!existing) {
            return { error: tm('errNoDeckAppend') }
          }
          const opened = existing.opened
          const beforeCount = opened.deck.slides.length
          // Push an undo snapshot: appending is an ordinary edit, ⌘Z should return to the
          // pre-append state (previously the undoStack was simply cleared, making all of the
          // user's prior manual edits non-undoable — inconsistent with replace_at behavior)
          pushHistory(existing)
          let merged = 0
          let lastErr: string | undefined
          for (const html of pagesHtml) {
            try {
              const one = await readCloudPage(html)
              const slide = await mergeSlideFromPptx(opened, one.bytes)
              if (slide) {
                promoteSlideBackground(slide, opened.deck.size)
                merged += 1
              } else lastErr = tm('errMergeFailed')
            } catch (pageErr) {
              lastErr = pageErr instanceof Error ? pageErr.message : String(pageErr)
            }
          }
          if (merged === 0) {
            existing.undoStack.pop() // Nothing happened, pop the just-pushed snapshot
            return { error: tm('errAppendFailed', { reason: lastErr ?? tm('errUnknown') }) }
          }
          existing.fitWidthPx = fitWidthPx
          // Save the draft: persist the current complete deck
          const bytes = await savePptx(opened)
          await saveDraftAfterGenerate(e.sender, existing, bytes, 'append', deckName)
          // Draft now matches memory: reopen from the output bytes to clear dirty (same as
          // slides:save) — otherwise pure AI generation (per-page append merges mark
          // structureDirty) would trigger the close confirmation even without edits
          if (existing.path) {
            existing.opened = await openPptx(bytes)
            existing.metaDirty = false
          }
          return {
            path: existing.path,
            slides: buildAllRenderSlides(existing.opened, fitWidthPx),
            size: { cx: existing.opened.deck.size.cx, cy: existing.opened.deck.size.cy },
            defaultFont: deckDefaultFont(existing.opened),
            appendedFrom: beforeCount,
            ...(lastErr && merged < pagesHtml.length
              ? { fallbackReason: tm('errPartialAppend', { reason: lastErr }) }
              : {}),
          }
        }

        // Redo one page in place: single-page HTML -> single-page pptx -> merge at the end ->
        // moveSlide into position -> delete the old page. Conversion happens first (deck
        // untouched); the mutation phase takes one undo snapshot overall, so ⌘Z rolls back to the
        // old page.
        if (mode === 'replace_at') {
          const existing = sessions.get(e.sender.id)
          if (!existing) {
            return { error: tm('errNoDeckReplace') }
          }
          const opened = existing.opened
          const total = opened.deck.slides.length
          if (atIndex == null || !Number.isInteger(atIndex) || atIndex < 0 || atIndex >= total) {
            return { error: tm('errIndexRange', { max: total - 1 }) }
          }
          const html = pagesHtml[0]
          if (!html || pagesHtml.length !== 1) {
            return { error: tm('errReplaceNeedsOne') }
          }
          const one = await readCloudPage(html)
          pushHistory(existing)
          const rollback = () => {
            const snap = existing.undoStack.pop()
            if (snap) restoreSnapshot(existing, snap)
          }
          const merged = await mergeSlideFromPptx(opened, one.bytes)
          if (!merged) {
            rollback()
            return { error: tm('errMergeFailed') }
          }
          promoteSlideBackground(merged, opened.deck.size)
          // The new page is at the end (index=total); after moving to atIndex the old page gets pushed to atIndex+1, delete it
          if (!moveSlide(opened, total, atIndex) || !deleteSlide(opened, atIndex + 1)) {
            rollback()
            return { error: tm('errReplaceFailed') }
          }
          existing.fitWidthPx = fitWidthPx
          const bytes = await savePptx(opened)
          await saveDraftAfterGenerate(e.sender, existing, bytes, 'append', deckName)
          if (existing.path) {
            existing.opened = await openPptx(bytes)
            existing.metaDirty = false
          }
          return {
            path: existing.path,
            slides: buildAllRenderSlides(existing.opened, fitWidthPx),
            size: { cx: existing.opened.deck.size.cx, cy: existing.opened.deck.size.cy },
            defaultFont: deckDefaultFont(existing.opened),
            replacedIndex: atIndex,
          }
        }

        // Insert one page at atIndex (later pages shift back): used to regenerate a failed middle
        // page from generate_deck and put it back in place. Same mechanism as replace_at (merge at
        // the end -> moveSlide into position) but without deleting an old page.
        if (mode === 'insert_at') {
          const existing = sessions.get(e.sender.id)
          if (!existing) {
            return { error: tm('errNoDeckInsert') }
          }
          const opened = existing.opened
          const total = opened.deck.slides.length
          if (atIndex == null || !Number.isInteger(atIndex) || atIndex < 0 || atIndex > total) {
            return { error: tm('errIndexRange', { max: total }) }
          }
          const html = pagesHtml[0]
          if (!html || pagesHtml.length !== 1) {
            return { error: tm('errInsertNeedsOne') }
          }
          const one = await readCloudPage(html)
          pushHistory(existing)
          const rollback = () => {
            const snap = existing.undoStack.pop()
            if (snap) restoreSnapshot(existing, snap)
          }
          const merged = await mergeSlideFromPptx(opened, one.bytes)
          if (!merged) {
            rollback()
            return { error: tm('errMergeFailed') }
          }
          promoteSlideBackground(merged, opened.deck.size)
          // The new page is at the end (index=total); with atIndex=total it belongs at the end anyway, no move needed
          if (atIndex < total && !moveSlide(opened, total, atIndex)) {
            rollback()
            return { error: tm('errInsertFailed') }
          }
          existing.fitWidthPx = fitWidthPx
          const bytes = await savePptx(opened)
          await saveDraftAfterGenerate(e.sender, existing, bytes, 'append', deckName)
          if (existing.path) {
            existing.opened = await openPptx(bytes)
            existing.metaDirty = false
          }
          return {
            path: existing.path,
            slides: buildAllRenderSlides(existing.opened, fitWidthPx),
            size: { cx: existing.opened.deck.size.cx, cy: existing.opened.deck.size.cy },
            defaultFont: deckDefaultFont(existing.opened),
            insertedIndex: atIndex,
          }
        }

        // replace mode: assemble the whole batch into one multi-page pptx as the new deck base.
        const { bytes } = await assembleDeck()
        const opened = await openPptx(bytes)
        // With per-page conversion + merging, stored PageVisualData is no longer needed; append reads the opened deck directly.
        const replaceSession: Session = {
          path: '',
          opened,
          fitWidthPx,
          undoStack: [],
          redoStack: [],
          htmlPages: null,
        }
        carryHistoryForReplacement(sessions.get(e.sender.id), replaceSession)
        sessions.set(e.sender.id, replaceSession)
        // Save the draft: await completion so the real path is returned; on failure degrade silently (session.path stays '')
        await saveDraftAfterGenerate(e.sender, replaceSession, bytes, 'replace', deckName)
        return {
          path: replaceSession.path,
          slides: buildAllRenderSlides(opened, fitWidthPx),
          size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy },
          defaultFont: deckDefaultFont(opened),
        }
      } catch (err) {
        return { error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle('slides:new-blank', async (e, fitWidthPx: number): Promise<OpenResult> => {
    const opened = await openPptx(await createBlankPptx())
    sessions.set(e.sender.id, { path: '', opened, fitWidthPx, undoStack: [], redoStack: [] })
    scheduleHistoryNotify(sessions.get(e.sender.id)!)
    return {
      path: '',
      slides: buildAllRenderSlides(opened, fitWidthPx),
      size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy },
      defaultFont: deckDefaultFont(opened),
    }
  })

  ipcMain.handle('slides:add-element', (e, op: AddElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const paragraphs: Paragraph[] | undefined = op.paragraphs?.length
      ? (op.paragraphs as Paragraph[])
      : op.text
        ? op.text.split('\n').map((line) => ({ runs: [{ text: line }] }))
        : undefined
    const el = addElement(slide, {
      kind: op.kind,
      offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
      ...(paragraphs ? { paragraphs } : {}),
      ...(op.fillColor ? { fillColor: op.fillColor } : {}),
      ...(op.stroke
        ? {
            stroke: {
              color: op.stroke.color,
              widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
            },
          }
        : {}),
    })
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: el.id } : null
  })

  ipcMain.handle('slides:delete-element', (e, op: DeleteElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    if (!slide.elements.some((x) => x.id === op.sourceId)) return null
    pushHistory(session)
    if (!deleteElement(slide, op.sourceId)) return null
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:edit-stroke', (e, op: EditStrokeOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    if (op.groupId) {
      pushHistory(session)
      const stroke = op.stroke
        ? {
            color: op.stroke.color,
            widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
            ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
          }
        : null
      if (!editGroupChildStroke(slide, op.groupId, op.sourceId, stroke)) {
        session.undoStack.pop()
        return null
      }
      return rebuildSlide(session, op.slideIndex)
    }
    // Unlike findEl, pictures are strokable too (picture border)
    const el = slide.elements.find(
      (x) =>
        x.id === op.sourceId && (x.type === 'text' || x.type === 'shape' || x.type === 'picture'),
    ) as TextElement | undefined
    if (!el) return null
    pushHistory(session)
    el.stroke = op.stroke
      ? {
          fill: { type: 'solid', color: op.stroke.color },
          width: Math.round(op.stroke.widthPt * EMU_PER_PT),
          ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
        }
      : undefined
    el.dirtyStroke = true
    return rebuildSlide(session, op.slideIndex)
  })

  // Mirror elements across their own axis: flipH/flipV is the only way to
  // point an arrow the other way — rotation cannot express a single-axis mirror
  ipcMain.handle('slides:flip-elements', (e, op: FlipElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    // Any element with an a:xfrm can flip (picture/shape/text/group) — findEl's
    // text/shape filter would silently drop pictures
    const targets = op.sourceIds
      .map((id) =>
        op.groupId
          ? findGroupChild(slide, op.groupId, id)?.child
          : slide.elements.find((x) => x.id === id),
      )
      .filter((el): el is NonNullable<typeof el> => !!el && !!el.transform)
    if (targets.length === 0) return null
    pushHistory(session)
    for (const el of targets) {
      const t = el.transform
      // Rotation pivots on the flip-adjusted box origin, so toggling a flip on a
      // rotated element would move its visual center: origin + R(rot)·(flip-signed
      // half-extent) must stay put — shift the offset by the orbit difference.
      const orbit = () => {
        const rad = (((t.rot ?? 0) / 60000) * Math.PI) / 180
        const bx = t.flipH ? t.offset.cx : 0
        const by = t.flipV ? t.offset.cy : 0
        const vx = ((t.flipH ? -1 : 1) * t.offset.cx) / 2
        const vy = ((t.flipV ? -1 : 1) * t.offset.cy) / 2
        return {
          x: bx + vx * Math.cos(rad) - vy * Math.sin(rad),
          y: by + vx * Math.sin(rad) + vy * Math.cos(rad),
        }
      }
      const before = orbit()
      if (op.axis === 'h') t.flipH = !t.flipH
      else t.flipV = !t.flipV
      const after = orbit()
      t.offset.x += Math.round(before.x - after.x)
      t.offset.y += Math.round(before.y - after.y)
      el.dirtyTransform = true
    }
    updateConnectorsForMoved(
      slide,
      targets.map((el) => el.id),
    )
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:edit-picture-src-rect', (e, op: EditPictureSrcRectOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!editPictureSrcRect(slide, op.sourceId, op.srcRect)) {
      session.undoStack.pop()
      return null
    }
    // Crop confirm also shrinks the element frame to the crop frame — same history
    // push, so a single undo restores frame and crop together (no distorted middle state)
    if (op.boxPx && op.fitWidthPx) {
      const el = slide.elements.find((x) => x.id === op.sourceId)
      if (el) {
        const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
        const scale = op.fitWidthPx / baseWidthPx
        const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
        el.transform = {
          ...el.transform,
          offset: {
            x: toEmu(op.boxPx.x),
            y: toEmu(op.boxPx.y),
            cx: toEmu(op.boxPx.w),
            cy: toEmu(op.boxPx.h),
          },
        }
        el.dirtyTransform = true
        updateConnectorsForMoved(slide, [op.sourceId])
      }
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:group-elements', (e, op: GroupElementsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const result = groupElements(session.opened, op.slideIndex, op.sourceIds)
    if (!result) {
      session.undoStack.pop()
      return null
    }
    // groupElements already updated deck.slides[slideIndex] internally via materializeSlide
    const renderSlide = rebuildSlide(session, op.slideIndex)
    return renderSlide ? { slide: renderSlide, groupId: result.groupId } : null
  })

  ipcMain.handle('slides:ungroup-element', (e, op: UngroupElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const fresh = ungroupElement(session.opened, op.slideIndex, op.sourceId)
    if (!fresh) {
      session.undoStack.pop()
      return null
    }
    // ungroupElement already updated deck.slides[slideIndex] internally
    return rebuildSlide(session, op.slideIndex)
  })

  // Full-page "backdrop" rectangles: design templates often use a text-free solid rectangle
  // covering the whole page as background; changing only the page background would be hidden
  // behind them — so repaint such rectangles along with the background (solid/gradient), or
  // make them transparent when the new background is a picture.
  const repaintFullBleedBackdrops = (
    slide: Slide,
    size: { cx: number; cy: number },
    fill:
      | { type: 'solid'; color: string }
      | {
          type: 'gradient'
          stops: Array<{ pos: number; color: string }>
          angle?: number
          path?: 'circle'
        }
      | { type: 'none' },
  ) => {
    for (const el of slide.elements) {
      if (el.type !== 'shape' && el.type !== 'text') continue
      const shaped = el as TextElement
      const fillType = shaped.fill?.type
      if (fillType !== 'solid' && fillType !== 'gradient') continue
      if (shaped.text?.paragraphs.some((p) => p.runs.some((r) => r.text.trim()))) continue
      const { x, y, cx, cy } = el.transform.offset
      const coversX = x <= size.cx * 0.05 && x + cx >= size.cx * 0.95
      const coversY = y <= size.cy * 0.05 && y + cy >= size.cy * 0.95
      if (!coversX || !coversY) continue
      shaped.fill = fill
      shaped.dirtyFill = true
    }
  }

  ipcMain.handle('slides:edit-background', async (e, op: EditBackgroundOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slides = session.opened.deck.slides
    const targets = op.slideIndex === -1 ? slides : [slides[op.slideIndex]].filter(Boolean)
    if (targets.length === 0) return null

    if (op.kind === 'image') {
      let source: { bytes: Uint8Array; ext: string } | { mediaPath: string }
      if (op.pick !== false) {
        const r = await showOpenDialogWithMemory(dialog, dialogParent(), {
          title: tm('dlgInsertImage'),
          properties: ['openFile' as const],
          filters: [
            {
              name: tm('filterImages'),
              extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'],
            },
          ],
        })
        if (r.canceled || !r.filePaths[0]) return null
        const bytes = await readFile(r.filePaths[0])
        source = {
          bytes: new Uint8Array(bytes),
          ext: r.filePaths[0].split('.').pop()!.toLowerCase(),
        }
      } else {
        // Reuse an already-landed background image (mode change / apply-to-all)
        const src = slides[op.sourceSlideIndex ?? op.slideIndex]
        if (src?.background?.type !== 'image') return null
        source = { mediaPath: src.background.mediaRef }
      }
      pushHistory(session)
      // The picked bytes land as one media part; further slides only add a rel to it
      let landed: string | null = null
      for (const s of targets) {
        const used = setSlideBackgroundImage(
          session.opened,
          s!,
          landed ? { mediaPath: landed } : source,
          op.mode === 'tile',
        )
        if (used) {
          landed = used
          repaintFullBleedBackdrops(s!, session.opened.deck.size, { type: 'none' })
        }
      }
      if (!landed) {
        session.undoStack.pop()
        return null
      }
    } else if (op.kind === 'solid') {
      pushHistory(session)
      for (const s of targets) {
        setSlideBackground(s!, op.color)
        repaintFullBleedBackdrops(s!, session.opened.deck.size, {
          type: 'solid',
          color: op.color,
        })
      }
    } else if (op.kind === 'gradient') {
      pushHistory(session)
      const stops = [
        { pos: 0, color: op.from },
        { pos: 1, color: op.to },
      ]
      const angle = Math.round((op.angleDeg ?? 0) * 60000)
      const patch = { stops, ...(op.radial ? { radial: true } : { angle }) }
      for (const s of targets) {
        setSlideBackground(s!, patch)
        repaintFullBleedBackdrops(s!, session.opened.deck.size, {
          type: 'gradient',
          stops,
          ...(op.radial ? { path: 'circle' as const } : { angle }),
        })
      }
    } else if (op.kind === 'reset') {
      pushHistory(session)
      for (const s of targets) resetSlideBackground(session.opened, s!)
    } else {
      pushHistory(session)
      for (const s of targets) setSlideBgGraphicsHidden(session.opened, s!, op.hidden)
    }
    session.fitWidthPx = op.fitWidthPx
    return buildAllRenderSlides(session.opened, op.fitWidthPx)
  })

  ipcMain.handle(
    'slides:edit-image-fill',
    async (e, op: { slideIndex: number; sourceId: string }) => {
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const slide = session.opened.deck.slides[op.slideIndex]
      if (!slide) return null
      const parent = dialogParent()
      const options = {
        title: tm('dlgInsertImage'),
        properties: ['openFile' as const],
        filters: [
          {
            name: tm('filterImages'),
            extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'],
          },
        ],
      }
      const r = await showOpenDialogWithMemory(dialog, parent, options)
      if (r.canceled || !r.filePaths[0]) return null
      const bytes = await readFile(r.filePaths[0])
      const ext = r.filePaths[0].split('.').pop()!.toLowerCase()
      pushHistory(session)
      if (!setElementImageFill(session.opened, slide, op.sourceId, bytes, ext)) {
        session.undoStack.pop()
        return null
      }
      return rebuildSlide(session, op.slideIndex)
    },
  )

  ipcMain.handle('slides:insert-image', async (e, slideIndex: number, fitWidthPx: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[slideIndex]
    if (!slide) return null
    const parent = dialogParent()
    const options = {
      title: tm('dlgInsertImage'),
      properties: ['openFile' as const],
      filters: [
        {
          name: tm('filterImages'),
          extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'],
        },
      ],
    }
    const r = await showOpenDialogWithMemory(dialog, parent, options)
    if (r.canceled || !r.filePaths[0]) return null
    const filePath = r.filePaths[0]
    const bytes = await readFile(filePath)
    const ext = filePath.split('.').pop()!.toLowerCase()

    // Scale proportionally to at most half the page width/height, centered
    const deckSize = session.opened.deck.size
    let natural = { width: 4, height: 3 }
    if (ext === 'tif' || ext === 'tiff') {
      const decoded = tiffToPng(new Uint8Array(bytes))
      if (decoded) natural = { width: decoded.width, height: decoded.height }
    } else {
      const img = nativeImage.createFromPath(filePath)
      if (!img.isEmpty()) natural = img.getSize()
    }
    const maxW = deckSize.cx / 2
    const maxH = deckSize.cy / 2
    const scale = Math.min(maxW / natural.width, maxH / natural.height)
    const cx = Math.round(natural.width * scale)
    const cy = Math.round(natural.height * scale)
    const offset = {
      x: Math.round((deckSize.cx - cx) / 2),
      y: Math.round((deckSize.cy - cy) / 2),
      cx,
      cy,
    }

    pushHistory(session)
    const el = addPicture(session.opened, slide, { bytes: new Uint8Array(bytes), ext, offset })
    if (!el) {
      session.undoStack.pop()
      return { error: 'unsupported' as const, ext }
    }
    session.fitWidthPx = fitWidthPx
    const rebuilt = rebuildSlide(session, slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: el.id } : null
  })

  ipcMain.handle('slides:edit-fill', (e, op: EditFillOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    if (op.groupId) {
      pushHistory(session)
      const fill =
        typeof op.fill === 'string'
          ? op.fill
          : {
              stops: [
                { pos: 0, color: op.fill.gradient.from },
                { pos: 1, color: op.fill.gradient.to },
              ],
              ...(op.fill.gradient.radial
                ? { radial: true }
                : { angle: Math.round((op.fill.gradient.angleDeg ?? 0) * 60000) }),
            }
      if (!editGroupChildFill(slide, op.groupId, op.sourceId, fill)) {
        session.undoStack.pop()
        return null
      }
      return rebuildSlide(session, op.slideIndex)
    }
    const el = findEl(slide, op.sourceId)
    if (!el) return null
    pushHistory(session)
    if (typeof op.fill === 'string') {
      el.fill = op.fill === 'none' ? { type: 'none' } : { type: 'solid', color: op.fill }
    } else {
      const g = op.fill.gradient
      el.fill = {
        type: 'gradient',
        stops: [
          { pos: 0, color: g.from },
          { pos: 1, color: g.to },
        ],
        ...(g.radial
          ? { path: 'circle' as const }
          : { angle: Math.round((g.angleDeg ?? 0) * 60000) }),
      }
    }
    el.dirtyFill = true
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:add-slide', (e, op: AddSlideOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const slide = duplicateSlide(session.opened, op.sourceIndex, { clearText: !!op.clearText })
    if (!slide) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: op.sourceIndex + 1,
    }
  })

  // App-wide, so a slide copied in one tab can be pasted into another deck.
  ipcMain.handle('slides:copy-slide', (e, slideIndex: number, pngBase64?: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    const bundle = copySlide(session.opened, slideIndex)
    if (!bundle) return false
    slideClipboard = { bundle, ...(pngBase64 ? { png: pngBase64 } : {}) }
    // Marker so plain ⌘V knows the latest copy was a slide (element copies / external copies overwrite it)
    clipboard.writeBuffer('io.genoffice.slides.slide', Buffer.from('1'))
    return true
  })

  ipcMain.handle('slides:has-slide-clipboard', () => slideClipboard !== null)

  const performSlidePaste = (
    session: Session,
    op: PasteSlideOp,
  ): { slides: RenderSlide[]; index: number; sourceId?: string } | null => {
    if (!slideClipboard) return null
    if (op.mode === 'picture') {
      const { deck } = session.opened
      const anchorIndex = Math.min(Math.max(op.afterIndex, 0), deck.slides.length - 1)
      const slide = deck.slides[anchorIndex]
      if (!slide || !slideClipboard.png) return null
      const el = addPicture(session.opened, slide, {
        bytes: new Uint8Array(Buffer.from(slideClipboard.png, 'base64')),
        ext: 'png',
        offset: { x: 0, y: 0, cx: deck.size.cx, cy: deck.size.cy },
      })
      if (!el) return null
      session.fitWidthPx = op.fitWidthPx
      return {
        slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
        index: anchorIndex,
        sourceId: el.id,
      }
    }
    const slide = pasteSlide(session.opened, op.afterIndex, slideClipboard.bundle, {
      keepSourceFormatting: op.mode === 'source',
    })
    if (!slide) return null
    session.fitWidthPx = op.fitWidthPx
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: session.opened.deck.slides.indexOf(slide),
    }
  }

  ipcMain.handle('slides:paste-slide', (e, op: PasteSlideOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !slideClipboard) return null
    pushHistory(session)
    const r = performSlidePaste(session, op)
    if (!r) {
      session.undoStack.pop()
      return null
    }
    lastSlidePaste.set(e.sender.id, {
      afterIndex: op.afterIndex,
      undoLen: session.undoStack.length,
    })
    return r
  })

  // Paste-options floater: undo the just-completed paste and redo it with another
  // mode. Refused when anything (edits, ⌘Z) touched the deck in between.
  ipcMain.handle('slides:repaste-slide', (e, op: RepasteSlideOp) => {
    const session = sessions.get(e.sender.id)
    const rec = lastSlidePaste.get(e.sender.id)
    if (!session || !slideClipboard || !rec) return null
    if (session.undoStack.length !== rec.undoLen) return null
    const snap = session.undoStack.pop()
    if (!snap) return null
    restoreSnapshot(session, snap)
    pushHistory(session)
    const r = performSlidePaste(session, {
      afterIndex: rec.afterIndex,
      fitWidthPx: op.fitWidthPx,
      mode: op.mode,
    })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    rec.undoLen = session.undoStack.length
    return r
  })

  ipcMain.handle('slides:add-blank-slide', (e, op: AddBlankSlideOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const slide = insertBlankSlide(session.opened, op.sourceIndex)
    if (!slide) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: op.sourceIndex + 1,
    }
  })

  // 'builtin:<key>' virtual paths get injected into the package on first use
  const resolveLayoutPath = (session: Session, layoutPath?: string): string | undefined => {
    if (!layoutPath?.startsWith(BUILTIN_LAYOUT_PREFIX)) return layoutPath
    return (
      ensureBuiltinLayout(
        session.opened.archive,
        session.opened.deck.size,
        layoutPath.slice(BUILTIN_LAYOUT_PREFIX.length),
      ) ?? undefined
    )
  }

  ipcMain.handle('slides:add-slide-with-layout', (e, op: AddSlideWithLayoutOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const layoutPath = resolveLayoutPath(session, op.layoutPath)
    const slide = layoutPath
      ? insertSlideWithLayout(session.opened, op.sourceIndex, layoutPath)
      : null
    if (!slide) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: op.sourceIndex + 1,
    }
  })

  ipcMain.handle('slides:get-layouts', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const layouts = listSlideLayouts(session.opened.archive)
    // Decks whose own layouts carry no placeholders (AI-generated single blank layout)
    // get the built-in standard set, injected into the package on first use
    if (shouldOfferBuiltinLayouts(layouts)) {
      layouts.push(
        ...builtinLayoutInfos(session.opened.deck.size, new Set(layouts.map((l) => l.name))),
      )
    }
    return { layouts, size: { ...session.opened.deck.size } }
  })

  // ── Master edit view ───────────────────────────────────────────────
  // Exception to the fidelity rule: only parts the user actively changed in master view are
  // written back, using the same byte surgery as slides. Every commit writes the entry + fully
  // reparses all slides — inheritance takes effect immediately, and each undo snapshot's
  // (slides model, entries) pair stays self-consistent (rendering and file don't diverge after
  // undo).
  const buildMasterRenderSlide = (session: Session): RenderSlide | null => {
    const me = session.masterEdit
    if (!me) return null
    return buildRenderSlide(me.slide, session.opened.deck.size, {
      fitWidthPx: session.fitWidthPx,
      media: makeMediaResolver(session.opened),
      metrics: getFontMetrics(),
    })
  }

  const commitMasterEdit = (session: Session): void => {
    const me = session.masterEdit!
    session.opened.archive.entries.set(me.partPath, Buffer.from(patchSlideXml(me.slide), 'utf8'))
    for (let i = 0; i < session.opened.deck.slides.length; i++) materializeSlide(session.opened, i)
    session.metaDirty = true
  }

  ipcMain.handle('slides:master-enter', (e, fitWidthPx: number): MasterEnterResult | null => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    session.fitWidthPx = fitWidthPx
    const items: MasterEnterResult['items'] = []
    for (const p of listMasterParts(session.opened.archive)) {
      const slide = parseMasterPart(session.opened.archive, p.partPath)
      if (!slide) continue
      const rendered = buildRenderSlide(slide, session.opened.deck.size, {
        fitWidthPx,
        media: makeMediaResolver(session.opened),
        metrics: getFontMetrics(),
      })
      items.push({ partPath: p.partPath, kind: p.kind, name: p.name, slide: rendered })
      if (!session.masterEdit) session.masterEdit = { partPath: p.partPath, slide }
    }
    return items.length ? { items } : null
  })

  ipcMain.handle('slides:master-open', (e, partPath: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = parseMasterPart(session.opened.archive, partPath)
    if (!slide) return null
    session.masterEdit = { partPath, slide }
    return buildMasterRenderSlide(session)
  })

  ipcMain.handle('slides:master-close', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    session.masterEdit = null
    // Edits were materialized one by one; here we only fetch the full render tree
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:master-edit-text', (e, op: MasterEditTextOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const el = findEl(me.slide, op.sourceId)
    if (!el?.text) return null
    pushHistory(session)
    el.text.paragraphs = applyEditParagraphs(el.text.paragraphs, op.paragraphs)
    el.dirty = true
    commitMasterEdit(session)
    return buildMasterRenderSlide(session)
  })

  ipcMain.handle('slides:master-edit-transform', (e, op: MasterEditTransformOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const el = me.slide.elements.find((x) => x.id === op.sourceId)
    if (!el) return null
    if (op.preview) {
      if (!session.transformPreview) {
        pushHistory(session)
        session.transformPreview = true
      }
    } else if (session.transformPreview) {
      session.transformPreview = false
    } else {
      pushHistory(session)
    }
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    el.transform = {
      ...el.transform,
      offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
      rot: Math.round(op.rotationDeg * 60000),
    }
    el.dirtyTransform = true
    // Previews are not persisted (only the final commit at drag end writes the entry + full reparse)
    if (!op.preview) commitMasterEdit(session)
    return buildMasterRenderSlide(session)
  })

  ipcMain.handle('slides:master-edit-fill', (e, op: MasterEditFillOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const el = findEl(me.slide, op.sourceId)
    if (!el) return null
    pushHistory(session)
    if (typeof op.fill === 'string') {
      el.fill = op.fill === 'none' ? { type: 'none' } : { type: 'solid', color: op.fill }
    } else {
      const g = op.fill.gradient
      el.fill = {
        type: 'gradient',
        stops: [
          { pos: 0, color: g.from },
          { pos: 1, color: g.to },
        ],
        ...(g.radial
          ? { path: 'circle' as const }
          : { angle: Math.round((g.angleDeg ?? 0) * 60000) }),
      }
    }
    el.dirtyFill = true
    commitMasterEdit(session)
    return buildMasterRenderSlide(session)
  })

  ipcMain.handle('slides:master-edit-stroke', (e, op: MasterEditStrokeOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const el = findEl(me.slide, op.sourceId)
    if (!el) return null
    pushHistory(session)
    el.stroke = op.stroke
      ? {
          fill: { type: 'solid', color: op.stroke.color },
          width: Math.round(op.stroke.widthPt * EMU_PER_PT),
        }
      : undefined
    el.dirtyStroke = true
    commitMasterEdit(session)
    return buildMasterRenderSlide(session)
  })

  ipcMain.handle('slides:master-delete-element', (e, op: MasterDeleteElementOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    if (!me.slide.elements.some((x) => x.id === op.sourceId)) return null
    pushHistory(session)
    if (!deleteElement(me.slide, op.sourceId)) {
      session.undoStack.pop()
      return null
    }
    commitMasterEdit(session)
    return buildMasterRenderSlide(session)
  })

  ipcMain.handle('slides:edit-picture-opacity', (e, op: EditPictureOpacityOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!setPictureOpacity(slide, op.sourceId, op.opacity)) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:set-slide-size', (e, op: SetSlideSizeOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    if (!setSlideSize(session.opened, op.cx, op.cy)) {
      session.undoStack.pop()
      return null
    }
    session.metaDirty = true
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:get-slide-size', (e) => {
    const session = sessions.get(e.sender.id)
    return session ? { ...session.opened.deck.size } : null
  })

  ipcMain.handle('slides:set-slide-layout', (e, op: SetSlideLayoutOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const layoutPath = resolveLayoutPath(session, op.layoutPath)
    const r = layoutPath
      ? setSlideLayout(session.opened, op.slideIndex, layoutPath)
      : op.layoutPath
        ? null
        : resetSlideLayout(session.opened, op.slideIndex)
    if (!r) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:find-replace', (e, op: FindReplaceOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const { count } = replaceAllInDeck(session.opened.deck, op.find, op.replace, {
      matchCase: op.matchCase,
      firstOnly: op.firstOnly,
      slideIndex: op.slideIndex,
      elementId: op.elementId,
    })
    if (!count) {
      session.undoStack.pop()
      return { count: 0, slides: null }
    }
    return { count, slides: buildAllRenderSlides(session.opened, session.fitWidthPx) }
  })

  ipcMain.handle('slides:delete-slide', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    if (!deleteSlide(session.opened, slideIndex)) {
      session.undoStack.pop()
      return null
    }
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:edit-table-cell', (e, op: EditTableCellOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!editTableCellText(slide, op.sourceId, op.row, op.col, op.paragraphs as Paragraph[])) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:table-merge', (e, op: TableMergeIpcOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const r = mergeTableCells(session.opened, op.slideIndex, op.sourceId, {
      kind: op.kind,
      row: op.row,
      col: op.col,
    })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.elementId } : null
  })

  ipcMain.handle('slides:table-structure', (e, op: TableStructureIpcOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const r = editTableStructure(session.opened, op.slideIndex, op.sourceId, {
      kind: op.kind,
      index: op.index,
      ...(op.before ? { before: true } : {}),
    } as TableStructureOp)
    if (!r) {
      session.undoStack.pop()
      return null
    }
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.elementId } : null
  })

  ipcMain.handle('slides:set-table-row-height', (e, op: SetTableRowHeightOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    pushHistory(session)
    if (!setTableRowHeight(slide, op.sourceId, op.row, (op.hPx / scale) * EMU_PER_PX_96)) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:set-table-cell-anchor', (e, op: SetTableCellAnchorOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!setTableCellAnchor(slide, op.sourceId, op.row, op.col, op.anchor)) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:set-table-col-width', (e, op: SetTableColWidthOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    pushHistory(session)
    if (!setTableColWidth(slide, op.sourceId, op.col, (op.wPx / scale) * EMU_PER_PX_96)) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:edit-table-style', (e, op: EditTableStyleOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    // A reparse regenerates element ids: look up the new id by element index; the renderer uses it to keep the selection
    const elIdx = slide.elements.findIndex((el) => el.id === op.sourceId)
    pushHistory(session)
    // Parse op -> TableStyleEdit
    let edit: TableStyleEdit
    if (op.styleName && TABLE_STYLE_PRESETS[op.styleName]) {
      const preset = TABLE_STYLE_PRESETS[op.styleName]!
      // Inject fixed-color presets' style definitions into tableStyles.xml (built-in GUIDs track theme colors, so colors would drift)
      if (preset.styleId && preset.styleDefXml) {
        ensureTableStylePart(session.opened, preset.styleId, preset.styleDefXml)
      }
      // Applying a style-gallery preset in PowerPoint clears cells' direct fills/borders; otherwise direct formatting hides the style
      edit = {
        tblPrXml: preset.tblPrXml,
        clearDirectFormatting: true,
        // Grid-style presets use direct borders (the style mechanism only has inner lines and cannot draw the outer frame)
        ...(preset.border
          ? {
              borderPreset: 'all' as const,
              borderColor: preset.border.color,
              borderWidthEmu: preset.border.widthEmu,
            }
          : {}),
      }
    } else {
      const borderColor = op.borderColor ?? undefined
      const borderWidthEmu =
        op.borderWidthPt != null ? Math.round(op.borderWidthPt * EMU_PER_PT) : undefined
      edit = {
        ...(op.firstRow !== undefined ? { firstRow: op.firstRow } : {}),
        ...(op.bandRow !== undefined ? { bandRow: op.bandRow } : {}),
        ...(op.shadingColor !== undefined ? { shadingColor: op.shadingColor } : {}),
        ...(op.borderPreset !== undefined ? { borderPreset: op.borderPreset } : {}),
        ...(borderColor !== undefined ? { borderColor } : {}),
        ...(borderWidthEmu !== undefined ? { borderWidthEmu } : {}),
        ...(op.cells ? { cells: op.cells } : {}),
      }
    }
    if (!editTableStyle(slide, op.sourceId, edit)) {
      session.undoStack.pop()
      return null
    }
    // The patch is written on anchor.originalXml; a materialize reparse is needed before it shows in the render model
    const rebuilt = rebuildSlideWithReparse(session, op.slideIndex)
    if (!rebuilt) return null
    const newId = session.opened.deck.slides[op.slideIndex]?.elements[elIdx]?.id ?? null
    return { slide: rebuilt, sourceId: newId }
  })

  ipcMain.handle('slides:edit-chart', async (e, op: EditChartOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    // A reparse regenerates element ids: look up the new id by element index; the renderer uses it to keep the selection
    const elIdx = slide.elements.findIndex((el) => el.id === op.sourceId)
    // Confirm before the first edit of a chart from an imported file: editing rebuilds it from the template,
    // and unmodeled fine-grained formatting (number formats/trendlines/error bars/per-point styles) is lost
    const chartEl = slide.elements[elIdx] as { type?: string; descr?: string } | undefined
    if (chartEl?.type === 'chart' && chartEl.descr !== 'aislides-chart') {
      const parent = dialogParent()
      const options = {
        type: 'warning' as const,
        buttons: [tm('chartSimplifyOk'), tm('btnCancel')],
        defaultId: 0,
        cancelId: 1,
        message: tm('chartSimplifyTitle'),
        detail: tm('chartSimplifyBody'),
      }
      const r = parent
        ? await dialog.showMessageBox(parent, options)
        : await dialog.showMessageBox(options)
      if (r.response !== 0) return null
    }
    pushHistory(session)
    // Mark aislides-chart on first edit (the conversion itself is lossless; no re-prompt after one confirmation)
    markChartEditable(slide, op.sourceId)
    const patch: Parameters<typeof editChartElement>[3] = {
      ...(op.kind ? { kind: op.kind === 'barH' ? 'bar' : op.kind } : {}),
      ...(op.kind === 'barH' ? { barDir: 'bar' as const } : {}),
      ...(op.categories ? { categories: op.categories } : {}),
      ...(op.series ? { series: op.series } : {}),
      ...(op.title !== undefined ? { title: op.title } : {}),
      ...(op.colorScheme
        ? {
            colorScheme:
              chartColorSchemes(session.opened).find((s) => s.key === op.colorScheme)?.colors ??
              CHART_COLOR_SCHEMES[op.colorScheme],
          }
        : {}),
      ...(op.legendPos ? { legendPos: op.legendPos } : {}),
      ...(op.dataLabels !== undefined ? { dataLabels: op.dataLabels } : {}),
      ...(op.gridlines !== undefined ? { gridlines: op.gridlines } : {}),
      ...(op.catAxisTitle !== undefined ? { catAxisTitle: op.catAxisTitle } : {}),
      ...(op.valAxisTitle !== undefined ? { valAxisTitle: op.valAxisTitle } : {}),
      ...(op.gapWidthPct !== undefined ? { gapWidthPct: op.gapWidthPct } : {}),
      ...(op.switchRowCol ? { switchRowCol: true } : {}),
      ...(op.pointColors ? { pointColors: op.pointColors } : {}),
    }
    if (!editChartElement(session.opened, op.slideIndex, op.sourceId, patch)) {
      session.undoStack.pop()
      return null
    }
    // The chart part XML is updated; reparse the whole page to refresh the model
    const rebuilt = rebuildSlideWithReparse(session, op.slideIndex)
    if (!rebuilt) return null
    const newId = session.opened.deck.slides[op.slideIndex]?.elements[elIdx]?.id ?? null
    return { slide: rebuilt, sourceId: newId }
  })

  ipcMain.handle('slides:chart-color-schemes', (e) => {
    const session = sessions.get(e.sender.id)
    return session ? chartColorSchemes(session.opened) : null
  })

  ipcMain.handle('slides:get-chart-data', (e, slideIndex: number, sourceId: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[slideIndex]
    if (!slide) return null
    return getChartElementData(slide, sourceId)
  })

  ipcMain.handle('slides:reorder-element', (e, op: ReorderElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!reorderElement(slide, op.sourceId, op.dir)) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle(
    'slides:set-text-anchor',
    (e, op: { slideIndex: number; sourceId: string; anchor: 'top' | 'middle' | 'bottom' }) => {
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const slide = session.opened.deck.slides[op.slideIndex]
      if (!slide) return null
      pushHistory(session)
      if (!setElementTextAnchor(slide, op.sourceId, op.anchor)) {
        session.undoStack.pop()
        return null
      }
      return rebuildSlide(session, op.slideIndex)
    },
  )

  ipcMain.handle('slides:clipboard-external', () => {
    // Our marker still present = the last copy came from this app -> use internal element paste
    // (on macOS custom formats don't appear in availableFormats, so check via readBuffer)
    const marker = (format: string) => {
      try {
        return clipboard.readBuffer(format).length > 0
      } catch {
        return false
      }
    }
    if (slideClipboard && marker('io.genoffice.slides.slide')) return { kind: 'slide' }
    if (marker('io.genoffice.slides.elements')) return { kind: 'internal' }
    const img = clipboard.readImage()
    if (!img.isEmpty()) return { kind: 'image', base64: img.toPNG().toString('base64'), ext: 'png' }
    const text = clipboard.readText()
    if (text.trim()) return { kind: 'text', text }
    return { kind: 'none' }
  })

  ipcMain.handle('slides:copy-elements', (e, op: CopyElementsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return 0
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return 0
    const items = op.sourceIds
      .map((id) => slide.elements.find((el) => el.id === id))
      .filter((el): el is NonNullable<typeof el> => !!el)
      .map((el) => copyElementData(session.opened, slide, el))
    if (items.length) {
      clipboards.set(e.sender.id, { items, pasteCount: 0 })
      // Write our marker to the OS clipboard: an external copy overwrites it, so at paste time it tells whether internal or external is newer
      clipboard.writeBuffer('io.genoffice.slides.elements', Buffer.from('1'))
    }
    return items.length
  })

  ipcMain.handle('slides:paste-elements', (e, op: PasteElementsOp) => {
    const session = sessions.get(e.sender.id)
    const clip = clipboards.get(e.sender.id)
    if (!session || !clip?.items.length) return null
    if (!session.opened.deck.slides[op.slideIndex]) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    // Cascading offset: each paste shifts another 16px relative to the original
    const shift = Math.round(((16 * (clip.pasteCount + 1)) / scale) * EMU_PER_PX_96)
    pushHistory(session)
    const r = pasteElements(session.opened, op.slideIndex, clip.items, { dx: shift, dy: shift })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    clip.pasteCount++
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceIds: r.elementIds } : null
  })

  // Duplicate in place (⌘D / Option+drag copy): does not touch the app clipboard; the caller supplies the offset
  ipcMain.handle('slides:duplicate-elements', (e, op: DuplicateElementsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const items = op.sourceIds
      .map((id) => slide.elements.find((el) => el.id === id))
      .filter((el): el is NonNullable<typeof el> => !!el)
      .map((el) => copyElementData(session.opened, slide, el))
    if (!items.length) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    pushHistory(session)
    const r = pasteElements(session.opened, op.slideIndex, items, {
      dx: toEmu(op.dxPx),
      dy: toEmu(op.dyPx),
    })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceIds: r.elementIds } : null
  })

  ipcMain.handle('slides:add-table', (e, op: AddTableOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    if (!session.opened.deck.slides[op.slideIndex]) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    pushHistory(session)
    const r = addTable(session.opened, op.slideIndex, {
      rows: op.rows,
      cols: op.cols,
      offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
    })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.elementId } : null
  })

  // Freehand ink stroke commit: one transparent PNG picture element per stroke (cNvPr name has
  // the aislides-ink prefix, descr stores the vector points as JSON); undo/save/thumbnails all
  // go through the existing picture-element pipeline.
  ipcMain.handle('slides:add-ink', (e, op: AddInkOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    pushHistory(session)
    const el = addPicture(session.opened, slide, {
      bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
      ext: 'png',
      offset: {
        x: toEmu(op.xPx),
        y: toEmu(op.yPx),
        cx: Math.max(1, toEmu(op.wPx)),
        cy: Math.max(1, toEmu(op.hPx)),
      },
      name: `aislides-ink ${Date.now().toString(36)}`,
      descr: op.payload,
    })
    if (!el) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: el.id } : null
  })

  // ── New insert capabilities: charts / SmartArt / icon bitmaps / audio-video / 3D / links / header-footer ──

  ipcMain.handle('slides:add-chart', (e, op: AddChartOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[op.slideIndex]) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    pushHistory(session)
    const r = addChart(session.opened, op.slideIndex, {
      kind: op.kind === 'barH' ? 'bar' : op.kind,
      ...(op.kind === 'barH' ? { barDir: 'bar' as const } : {}),
      ...(op.title ? { title: op.title } : {}),
      categories: op.categories,
      series: op.series,
      offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
    })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.elementId } : null
  })

  ipcMain.handle('slides:add-smartart', (e, op: AddSmartArtOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[op.slideIndex]) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    pushHistory(session)
    const r = addSmartArt(session.opened, op.slideIndex, {
      layout: op.layout,
      items: op.items,
      offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
    })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.elementId } : null
  })

  ipcMain.handle('slides:add-image-bytes', (e, op: AddImageBytesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    pushHistory(session)
    const el = addPicture(session.opened, slide, {
      bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
      ext: op.ext,
      offset: {
        x: toEmu(op.xPx),
        y: toEmu(op.yPx),
        cx: Math.max(1, toEmu(op.wPx)),
        cy: Math.max(1, toEmu(op.hPx)),
      },
      ...(op.name ? { name: op.name } : {}),
    })
    if (!el) {
      session.undoStack.pop()
      return { error: 'unsupported' as const, ext: op.ext }
    }
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: el.id } : null
  })

  ipcMain.handle('slides:replace-picture-bytes', (e, op: ReplacePictureBytesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    const ok = replacePictureBytes(
      session.opened,
      slide,
      op.sourceId,
      new Uint8Array(Buffer.from(op.base64, 'base64')),
      op.ext,
      op.keepSrcRect ? { keepSrcRect: true } : undefined,
    )
    if (!ok) {
      session.undoStack.pop()
      return { error: 'unsupported' as const, ext: op.ext }
    }
    return rebuildSlide(session, op.slideIndex)
  })

  // Show a dialog to pick video/audio and embed it. Video poster frame prefers the system thumbnail (QuickLook), falling back to a solid color on failure.
  ipcMain.handle(
    'slides:insert-media',
    async (e, slideIndex: number, kind: 'video' | 'audio', fitWidthPx: number) => {
      const session = sessions.get(e.sender.id)
      if (!session || !session.opened.deck.slides[slideIndex]) return null
      const parent = dialogParent()
      const filters =
        kind === 'video'
          ? [{ name: tm('filterVideo'), extensions: ['mp4', 'm4v', 'mov', 'webm', 'avi'] }]
          : [{ name: tm('filterAudio'), extensions: ['mp3', 'wav', 'm4a', 'aac', 'ogg'] }]
      const options = {
        title: kind === 'video' ? tm('dlgInsertVideo') : tm('dlgInsertAudio'),
        properties: ['openFile' as const],
        filters,
      }
      const r = await showOpenDialogWithMemory(dialog, parent, options)
      if (r.canceled || !r.filePaths[0]) return null
      const filePath = r.filePaths[0]
      const bytes = await readFile(filePath)
      const ext = filePath.split('.').pop()!.toLowerCase()
      const fileName = filePath.split('/').pop()!

      // Warn up front when in-app playback will be broken — AVI has no
      // Chromium demuxer at all; mp4/m4v/mov with e.g. AC-3/DTS audio plays silent.
      if (kind === 'video') {
        let detail: string | null = null
        if (ext === 'avi') detail = tm('mediaAviBody')
        else if (ext === 'mp4' || ext === 'm4v' || ext === 'mov') {
          const codec = unplayableAudioCodec(new Uint8Array(bytes))
          if (codec) detail = tm('mediaNoAudioBody', { codec })
        }
        if (detail) {
          const warn = {
            type: 'warning' as const,
            buttons: [tm('legacyPptOk')],
            message: tm('mediaUnsupportedTitle'),
            detail,
          }
          if (parent) await dialog.showMessageBox(parent, warn)
          else await dialog.showMessageBox(warn)
        }
      }

      let poster: { bytes: Uint8Array; ext: string } | undefined
      if (kind === 'video') {
        try {
          const thumb = await nativeImage.createThumbnailFromPath(filePath, {
            width: 960,
            height: 540,
          })
          if (!thumb.isEmpty()) poster = { bytes: new Uint8Array(thumb.toPNG()), ext: 'png' }
        } catch {
          /* Solid-color fallback */
        }
      }

      const deckSize = session.opened.deck.size
      const offset =
        kind === 'video'
          ? (() => {
              const cx = Math.round(deckSize.cx * 0.6)
              const cy = Math.round((cx * 9) / 16)
              return {
                x: Math.round((deckSize.cx - cx) / 2),
                y: Math.round((deckSize.cy - cy) / 2),
                cx,
                cy,
              }
            })()
          : (() => {
              const cx = Math.round(deckSize.cx * 0.24)
              const cy = Math.round(deckSize.cy * 0.09)
              return {
                x: Math.round((deckSize.cx - cx) / 2),
                y: Math.round((deckSize.cy - cy) / 2),
                cx,
                cy,
              }
            })()

      pushHistory(session)
      const added = addMedia(session.opened, slideIndex, {
        kind,
        bytes: new Uint8Array(bytes),
        ext,
        ...(poster ? { poster } : {}),
        offset,
        name: fileName,
      })
      if (!added) {
        session.undoStack.pop()
        return null
      }
      session.fitWidthPx = fitWidthPx
      const rebuilt = rebuildSlide(session, slideIndex)
      return rebuilt ? { slide: rebuilt, sourceId: added.elementId } : null
    },
  )

  // Double-click playback: read the media bytes of an audio/video element (embedded converts to dataUrl, external links return as-is)
  const AV_MIME: Record<string, string> = {
    mp4: 'video/mp4',
    m4v: 'video/mp4',
    // Chromium refuses to even load video/quicktime, but demuxes QuickTime bytes
    // fine through the ISO-BMFF path when served as video/mp4
    mov: 'video/mp4',
    webm: 'video/webm',
    avi: 'video/x-msvideo',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
    m4a: 'audio/mp4',
    aac: 'audio/aac',
    ogg: 'audio/ogg',
  }
  ipcMain.handle('slides:media-data', (e, slideIndex: number, sourceId: string) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!session || !slide) return null
    const el = slide.elements.find((x) => x.id === sourceId)
    if (!el || el.type !== 'picture') return null
    const media = (
      el as { media?: { kind: 'video' | 'audio'; target?: string; external?: boolean } }
    ).media
    if (!media?.target) return null
    if (media.external) return { kind: media.kind, dataUrl: media.target }
    const bytes = session.opened.archive.readBytes(media.target)
    if (!bytes) return null
    const ext = media.target.split('.').pop()?.toLowerCase() ?? ''
    const mime = AV_MIME[ext] ?? (media.kind === 'video' ? 'video/mp4' : 'audio/mpeg')
    return {
      kind: media.kind,
      dataUrl: `data:${mime};base64,${Buffer.from(bytes).toString('base64')}`,
    }
  })

  // Media recorded by the renderer (screen-recording webm): placed centered at 16:9
  ipcMain.handle('slides:add-media-bytes', (e, op: AddMediaBytesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[op.slideIndex]) return null
    const deckSize = session.opened.deck.size
    const cx = Math.round(deckSize.cx * 0.6)
    const cy = Math.round((cx * 9) / 16)
    pushHistory(session)
    const added = addMedia(session.opened, op.slideIndex, {
      kind: op.kind,
      bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
      ext: op.ext,
      offset: {
        x: Math.round((deckSize.cx - cx) / 2),
        y: Math.round((deckSize.cy - cy) / 2),
        cx,
        cy,
      },
      ...(op.name ? { name: op.name } : {}),
    })
    if (!added) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: added.elementId } : null
  })

  // 3D model (simplified): glb embed + poster placeholder image
  ipcMain.handle('slides:insert-model3d', async (e, slideIndex: number, fitWidthPx: number) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[slideIndex]) return null
    const parent = dialogParent()
    const options = {
      title: tm('dlgInsert3d'),
      properties: ['openFile' as const],
      filters: [{ name: tm('filter3d'), extensions: ['glb', 'gltf'] }],
    }
    const r = await showOpenDialogWithMemory(dialog, parent, options)
    if (r.canceled || !r.filePaths[0]) return null
    const filePath = r.filePaths[0]
    const bytes = await readFile(filePath)
    const ext = filePath.split('.').pop()!.toLowerCase()

    let poster: { bytes: Uint8Array; ext: string } | undefined
    try {
      const thumb = await nativeImage.createThumbnailFromPath(filePath, { width: 640, height: 640 })
      if (!thumb.isEmpty()) poster = { bytes: new Uint8Array(thumb.toPNG()), ext: 'png' }
    } catch {
      /* Dark-gray fallback */
    }

    const deckSize = session.opened.deck.size
    const cy = Math.round(deckSize.cy * 0.5)
    const cx = cy
    pushHistory(session)
    const added = addModel3d(session.opened, slideIndex, {
      bytes: new Uint8Array(bytes),
      ext,
      ...(poster ? { poster } : {}),
      offset: {
        x: Math.round((deckSize.cx - cx) / 2),
        y: Math.round((deckSize.cy - cy) / 2),
        cx,
        cy,
      },
      name: filePath.split('/').pop()!,
    })
    if (!added) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = fitWidthPx
    const rebuilt = rebuildSlide(session, slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: added.elementId } : null
  })

  ipcMain.handle('slides:set-link', (e, op: SetLinkOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[op.slideIndex]) return null
    pushHistory(session)
    const fresh = setElementLink(session.opened, op.slideIndex, op.sourceId, op.target)
    if (!fresh) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:get-link', (e, slideIndex: number, sourceId: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    return getElementLink(session.opened, slideIndex, sourceId)
  })

  ipcMain.handle('slides:get-slide-links', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return []
    return getSlideLinks(session.opened, slideIndex).map(({ elementId, target }) => ({
      sourceId: elementId,
      target,
    }))
  })

  ipcMain.handle('slides:get-run-links', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return []
    return getRunLinks(session.opened, slideIndex).map(({ elementId, ...rest }) => ({
      sourceId: elementId,
      ...rest,
    }))
  })

  ipcMain.handle('slides:apply-header-footer', (e, op: HeaderFooterOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const changed = applyHeaderFooter(session.opened, {
      footer: op.footer ?? null,
      slideNum: !!op.slideNum,
      date: op.date ?? null,
      ...(op.dateAuto ? { dateAuto: true } : {}),
    })
    if (!changed) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    return buildAllRenderSlides(session.opened, op.fitWidthPx)
  })

  ipcMain.handle('slides:get-header-footer', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return slide ? readHeaderFooter(slide) : { footer: null, slideNum: false, date: null }
  })

  // Apply a theme (Design tab theme gallery): rewrite theme*.xml colors/fonts (scheme-referenced
  // colors follow), and remap the deck's explicit srgbClr wholesale to the new theme palette
  // (real-world decks have almost entirely explicit colors, so swapping only the theme changes
  // nothing visually). Element resolved colors come from the parse-time inheritance chain, so
  // after the surgery the deck reparses in memory; undo snapshots roll back as usual.
  ipcMain.handle('slides:apply-theme', async (e, op: ApplyThemeOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const spec = {
      name: op.name,
      colors: op.colors,
      ...(op.majorFont ? { majorFont: op.majorFont } : {}),
      ...(op.minorFont ? { minorFont: op.minorFont } : {}),
    }
    try {
      // 1) Bake unsaved edits into the entries first: the color surgery edits entries
      //    directly, and dirty elements left for a later save would overwrite the surgery
      //    result with stale slices. In-memory (commitSaved/reparseDeck) instead of
      //    savePptx -> openPptx: the zip roundtrip's contiguous buffer fails on large decks
      commitSaved(session.opened)
      // 2) Pure entry surgery: theme parts + explicit color remapping
      const patched = applyThemeToArchive(session.opened, spec)
      const remapped = remapDeckColors(session.opened, spec)
      if (patched === 0 && remapped === 0) {
        session.undoStack.pop()
        return null
      }
      // 3) Reparse so every element's resolved colors/fonts refresh
      session.opened = reparseDeck(session.opened)
    } catch (err) {
      restoreSnapshot(session, session.undoStack.pop()!)
      return { error: err instanceof Error ? err.message : String(err) }
    }
    // Pages without any background definition fall back to the theme base color (so dark themes don't leave a white background)
    const lt1 = op.colors.lt1
    if (lt1) {
      for (const s of session.opened.deck.slides) {
        if (!s.background) setSlideBackground(s, `#${lt1.replace(/^#/, '')}`)
      }
    }
    // Reopening cleared element-level dirty; the session-level flag preserves the "unsaved" state (reset on save)
    session.metaDirty = true
    session.fitWidthPx = op.fitWidthPx
    return buildAllRenderSlides(session.opened, op.fitWidthPx)
  })

  ipcMain.handle('slides:set-transition', (e, op: SetTransitionOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    const slides = session.opened.deck.slides
    const targets = op.slideIndex === -1 ? slides : [slides[op.slideIndex]].filter(Boolean)
    if (targets.length === 0) return false
    pushHistory(session)
    for (const s of targets) setSlideTransition(s!, op.kind)
    return true
  })

  ipcMain.handle('slides:get-transition', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return slide ? getSlideTransition(slide) : 'none'
  })

  // Rehearsal timing save: batch-write each page's auto-advance time (<p:transition advTm>, ms)
  ipcMain.handle('slides:set-advance-times', (e, op: SetAdvanceTimesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    const slides = session.opened.deck.slides
    const targets = op.times.filter((t) => slides[t.slideIndex])
    if (targets.length === 0) return false
    pushHistory(session)
    for (const t of targets) setSlideAdvanceTime(slides[t.slideIndex]!, t.ms)
    return true
  })

  // ── Shape animations (<p:timing>; the spid <-> temporary element id mapping happens here) ──
  ipcMain.handle('slides:get-animations', (e, slideIndex: number): AnimationItem[] => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!slide) return []
    const bySpid = new Map<number, (typeof slide.elements)[number]>()
    for (const el of slide.elements) {
      const spid = elementSpid(el)
      if (spid != null && !bySpid.has(spid)) bySpid.set(spid, el)
    }
    const typeLabel: Record<string, string> = {
      text: tm('labelTextBox'),
      shape: tm('labelShape'),
      picture: tm('labelPicture'),
      group: tm('labelGroup'),
      table: tm('labelTable'),
      chart: tm('labelChart'),
      passthrough: tm('labelObject'),
    }
    const out: AnimationItem[] = []
    for (const a of getSlideAnimations(slide)) {
      const el = bySpid.get(a.spid)
      if (!el) continue // Leftover animations whose target shape was deleted are not echoed back
      out.push({
        sourceId: el.id,
        targetName: el.name || typeLabel[el.type] || tm('labelObject'),
        effect: a.effect,
        trigger: a.trigger,
        durationMs: a.durationMs,
        delayMs: a.delayMs,
        ...(a.motionPath != null ? { motionPath: a.motionPath } : {}),
        ...(a.paragraph != null ? { paragraph: a.paragraph } : {}),
      })
    }
    return out
  })

  // Pairing keys for Morph transitions: sourceId changes on every reparse, so match across pages by cNvPr id/name
  ipcMain.handle('slides:get-shape-keys', (e, slideIndex: number): ShapeKey[] => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!slide) return []
    return slide.elements.map((el) => ({
      sourceId: el.id,
      spid: elementSpid(el),
      name: el.name ?? '',
    }))
  })

  ipcMain.handle('slides:set-animations', (e, op: SetAnimationsOp) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[op.slideIndex]
    if (!session || !slide) return false
    const anims: SlideAnimation[] = []
    for (const it of op.items) {
      const el = slide.elements.find((x) => x.id === it.sourceId)
      const spid = el ? elementSpid(el) : null
      if (spid == null) continue
      anims.push({
        spid,
        effect: it.effect,
        trigger: it.trigger,
        durationMs: Math.max(0, Math.round(it.durationMs)),
        delayMs: Math.max(0, Math.round(it.delayMs)),
        ...(it.motionPath != null ? { motionPath: it.motionPath } : {}),
        ...(it.paragraph != null ? { paragraph: it.paragraph } : {}),
      })
    }
    pushHistory(session)
    setSlideAnimations(slide, anims)
    return true
  })

  ipcMain.handle('slides:set-hidden', (e, op: SetSlideHiddenOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    setSlideHidden(slide, op.hidden)
    return rebuildSlide(session, op.slideIndex)
  })

  // ── Section management: presentation.xml surgery, riding on snapshot undo and savePptx ──
  ipcMain.handle('slides:get-sections', (e) => {
    const session = sessions.get(e.sender.id)
    return session ? getSections(session.opened) : []
  })

  ipcMain.handle('slides:set-sections', (e, sections: SectionInfo[]) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    setSections(session.opened, sections)
    session.metaDirty = true
    return getSections(session.opened)
  })

  ipcMain.handle('slides:add-section', (e, op: AddSectionOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const r = addSection(session.opened, op.atSlideIndex, op.name)
    if (!r) {
      session.undoStack.pop()
      return null
    }
    session.metaDirty = true
    return r
  })

  ipcMain.handle('slides:rename-section', (e, op: RenameSectionOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const r = renameSection(session.opened, op.id, op.name)
    if (!r) {
      session.undoStack.pop()
      return null
    }
    session.metaDirty = true
    return r
  })

  ipcMain.handle('slides:remove-section', (e, op: RemoveSectionOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const r = removeSection(session.opened, op.id, { keepSlides: true })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    session.metaDirty = true
    return r
  })

  // Drag to reorder slides (sldIdLst + deck.slides + section membership); must send back the full RenderSlide set
  ipcMain.handle('slides:move-slide', (e, op: MoveSlideOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    if (!moveSlide(session.opened, op.fromIndex, op.toIndex)) {
      session.undoStack.pop()
      return null
    }
    session.metaDirty = true
    return {
      slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      sections: getSections(session.opened),
    }
  })

  // Moving a whole section changes slide order (sldIdLst + deck.slides); must send back the full RenderSlide set
  ipcMain.handle('slides:move-section', (e, op: MoveSectionOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const sections = moveSection(session.opened, op.id, op.dir)
    if (!sections) {
      session.undoStack.pop()
      return null
    }
    session.metaDirty = true
    return {
      slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      sections,
    }
  })

  // ── Speaker notes / comments (archive surgery, riding on snapshot undo and savePptx) ────
  ipcMain.handle('slides:get-notes', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return session && slide ? getSlideNotes(session.opened.archive, slide.path) : ''
  })

  ipcMain.handle('slides:set-notes', (e, op: SetNotesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[op.slideIndex]) return false
    pushHistory(session)
    const ok = setSlideNotes(session.opened, op.slideIndex, op.text)
    if (!ok) session.undoStack.pop()
    else session.metaDirty = true
    return ok
  })

  ipcMain.handle('slides:get-comments', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return session && slide ? getSlideComments(session.opened.archive, slide.path) : []
  })

  ipcMain.handle('slides:add-comment', (e, op: AddCommentOp) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[op.slideIndex]
    if (!session || !slide) return null
    pushHistory(session)
    const author = commentAuthorName()
    const added = addSlideComment(session.opened, op.slideIndex, { author, text: op.text })
    if (!added) {
      session.undoStack.pop()
      return null
    }
    session.metaDirty = true
    return getSlideComments(session.opened.archive, slide.path)
  })

  ipcMain.handle('slides:delete-comment', (e, op: DeleteCommentOp) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[op.slideIndex]
    if (!session || !slide) return null
    pushHistory(session)
    if (
      !deleteSlideComment(session.opened, op.slideIndex, { authorId: op.authorId, idx: op.idx })
    ) {
      session.undoStack.pop()
      return null
    }
    session.metaDirty = true
    return getSlideComments(session.opened.archive, slide.path)
  })

  // System clipboard while text-editing (menu commands are echoed back by the renderer per context)
  ipcMain.handle('slides:native-clipboard', (e, op: 'cut' | 'copy' | 'paste') => {
    if (op === 'cut') e.sender.cut()
    else if (op === 'copy') e.sender.copy()
    else e.sender.paste()
  })

  ipcMain.handle('slides:history-batch-begin', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    beginHistoryBatch(session)
    return true
  })

  ipcMain.handle('slides:history-batch-end', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const before = endHistoryBatch(session)
    return before ? registerAiSnapshot(session, before) : null
  })

  ipcMain.handle('slides:ai-snapshot-restore', (e, id: number) => {
    const session = sessions.get(e.sender.id)
    if (!session || session.masterEdit || session.historyBatch) return null
    if (!restoreAiSnapshot(session, id)) return null
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:undo', (e) => {
    const session = sessions.get(e.sender.id)
    // Undo disabled in master view: the masterEdit.slide model cannot roll back with snapshots (v1 trade-off; undoable after exiting)
    if (!session || session.masterEdit) return null
    settleStaleHistoryBatch(session)
    if (session.undoStack.length === 0) return null
    session.redoStack.push(takeSnapshot(session))
    restoreSnapshot(session, session.undoStack.pop()!)
    scheduleHistoryNotify(session)
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:redo', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session || session.masterEdit) return null
    settleStaleHistoryBatch(session)
    if (session.redoStack.length === 0) return null
    session.undoStack.push(takeSnapshot(session))
    restoreSnapshot(session, session.redoStack.pop()!)
    scheduleHistoryNotify(session)
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:is-dirty', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    return (
      !!session.metaDirty ||
      session.opened.deck.slides.some(
        (s) => s.structureDirty || s.elements.some((el) => el.dirty || el.dirtyTransform),
      )
    )
  })

  ipcMain.handle('slides:save', async (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return { ok: false, error: 'no file open' }
    // Untitled (new blank file): the first save lands silently in the drafts folder (Save As keeps its dialog)
    if (!session.path) {
      const draftsDir = getDraftsDir()
      if (!existsSync(draftsDir)) mkdirSync(draftsDir, { recursive: true })
      session.path = pickDraftPath(draftsDir, tm('untitledDeck'))
      await pushRecent(session.path)
      slidesOpenedHook?.(e.sender, session.path)
    }
    try {
      await savePptxToFile(session.opened, session.path)
      autosaveBackoff.delete(session.path)
      void rm(autosavePathFor(session.path), { force: true }).catch(() => {})
      dropUntitledRecovery(e.sender.id)
      // Bake the saved patches back into the in-memory model (clears dirty, syncs
      // anchor.originalXml with disk) — a full reopen would re-read and unzip the
      // whole package, doubling save latency on large decks. Element ids survive,
      // but the renderer still expects the render tree in the response.
      commitSaved(session.opened)
      session.metaDirty = false
      return {
        ok: true,
        path: session.path,
        slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      }
    } catch (err) {
      return { ok: false, error: String(err) }
    }
  })

  ipcMain.handle('slides:save-as', async (e, defaultName: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return { ok: false, error: 'no file open' }
    const parent = dialogParent()
    const options = {
      defaultPath: defaultName,
      filters: [{ name: 'PowerPoint', extensions: ['pptx'] }],
    }
    const r = await showSaveDialogWithMemory(dialog, parent, options, getDraftsDir())
    if (r.canceled || !r.filePath) return { ok: false }
    try {
      await savePptxToFile(session.opened, r.filePath)
      session.path = r.filePath
      autosaveBackoff.delete(r.filePath)
      dropUntitledRecovery(e.sender.id)
      await pushRecent(r.filePath)
      slidesOpenedHook?.(e.sender, r.filePath)
      commitSaved(session.opened)
      session.metaDirty = false
      return {
        ok: true,
        path: r.filePath,
        slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      }
    } catch (err) {
      return { ok: false, error: String(err) }
    }
  })

  // ── Export (PDF / images): the renderer renders hi-res PNGs with offscreen Konva; the main process handles dialogs/writing ──

  ipcMain.handle('slides:pick-export-dir', async () => {
    const parent = dialogParent()
    const options = {
      title: tm('dlgPickExportDir'),
      buttonLabel: tm('btnExport'),
      properties: ['openDirectory' as const, 'createDirectory' as const],
    }
    const r = await showOpenDialogWithMemory(dialog, parent, options)
    return r.canceled || !r.filePaths[0] ? null : r.filePaths[0]
  })

  ipcMain.handle(
    'slides:export-images',
    async (_e, op: ExportImagesOp): Promise<ExportImagesResult> => {
      try {
        // Zero-padding width follows the total page count (3 digits for ≥100 pages)
        const pad = op.pngsBase64.length >= 100 ? 3 : 2
        const paths: string[] = []
        for (let i = 0; i < op.pngsBase64.length; i++) {
          const p = join(op.dir, `${op.baseName}-${String(i + 1).padStart(pad, '0')}.png`)
          await writeFile(p, Buffer.from(op.pngsBase64[i], 'base64'))
          paths.push(p)
        }
        return { ok: true, paths }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  ipcMain.handle('slides:pick-export-pdf-path', async (_e, defaultName: string) => {
    const parent = dialogParent()
    const options = {
      title: tm('dlgExportPdf'),
      defaultPath: defaultName,
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    }
    const r = await showSaveDialogWithMemory(dialog, parent, options, getDraftsDir())
    return r.canceled || !r.filePath ? null : r.filePath
  })

  ipcMain.handle('slides:export-pdf', async (_e, op: ExportPdfOp): Promise<ExportPdfResult> => {
    // PDF page size: fixed 7.5in height, width by slide ratio (16:9 -> 13.333in, 4:3 -> 10in)
    const heightIn = 7.5
    const widthIn = Math.round((op.widthPx / op.heightPx) * heightIn * 1000) / 1000
    const html = `<!doctype html><html><head><meta charset="utf-8"><style>
@page { size: ${widthIn}in ${heightIn}in; margin: 0; }
html, body { margin: 0; padding: 0; }
.page { width: ${widthIn}in; height: ${heightIn}in; overflow: hidden; page-break-after: always; }
.page:last-child { page-break-after: auto; }
.page img { display: block; width: 100%; height: 100%; }
</style></head><body>${op.pngsBase64
      .map((b64) => `<div class="page"><img src="data:image/png;base64,${b64}"></div>`)
      .join('')}</body></html>`
    const win = new BrowserWindow({ show: false, webPreferences: { sandbox: true } })
    try {
      await win.loadURL('data:text/html;base64,' + Buffer.from(html, 'utf8').toString('base64'))
      // Wait for fonts and all images to decode before printing, avoiding blank pages
      await win.webContents.executeJavaScript(
        'Promise.all([document.fonts.ready, ...Array.from(document.images).map((i) => i.decode().catch(() => {}))])',
        true,
      )
      const pdf = await win.webContents.printToPDF({
        landscape: false, // The page size is already landscape (width > height); passing landscape would rotate a second time
        printBackground: true,
        pageSize: { width: widthIn, height: heightIn },
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        preferCSSPageSize: false,
      })
      await writeFile(op.filePath, pdf)
      return { ok: true, path: op.filePath }
    } catch (err) {
      return { ok: false, error: String(err) }
    } finally {
      win.destroy()
    }
  })

  ipcMain.handle(
    'slides:print',
    async (e, op: PrintSlidesOp): Promise<{ ok: boolean; error?: string }> => {
      // Page assembly is shared with the renderer's print-preview pane (print-html.ts)
      const html = buildPrintDocumentHtml({
        srcs: op.pngsBase64.map((b64) => `data:image/png;base64,${b64}`),
        ratio: op.widthPx / op.heightPx,
        layout: op.layout ?? 'full',
        ...(op.notes ? { notes: op.notes } : {}),
        ...(op.orientation ? { orientation: op.orientation } : {}),
        ...(op.frame ? { frame: true } : {}),
      })
      const owner = BrowserWindow.fromWebContents(e.sender) ?? dialogParent()
      const win = new BrowserWindow({
        show: false,
        ...(owner && !owner.isDestroyed() ? { parent: owner } : {}),
        ...(process.platform === 'win32'
          ? {
              width: 900,
              height: 700,
              autoHideMenuBar: true,
              closable: false,
              skipTaskbar: true,
            }
          : {}),
        webPreferences: { sandbox: true },
      })
      try {
        await win.loadURL('data:text/html;base64,' + Buffer.from(html, 'utf8').toString('base64'))
        await win.webContents.executeJavaScript(
          'Promise.all([document.fonts.ready, ...Array.from(document.images).map((i) => i.decode().catch(() => {}))])',
          true,
        )
        // Chromium attaches the native Windows print dialog to the window being printed.
        // If that owner is hidden, the dialog is hidden too and the layout buttons appear inert.
        if (process.platform === 'win32') {
          win.show()
          win.focus()
        }
        const result = await new Promise<{ success: boolean; failureReason: string }>((resolve) => {
          win.webContents.print(
            { silent: false, printBackground: true },
            (success, failureReason) => resolve({ success, failureReason }),
          )
        })
        if (!result.success) {
          // Canceling is a normal completion, not a print failure: ok=false without an
          // error keeps the renderer's print dialog (and its chosen options) open.
          if (result.failureReason === 'Print job canceled') return { ok: false }
          return { ok: false, error: result.failureReason }
        }
        return { ok: true }
      } catch (err) {
        return { ok: false, error: String(err) }
      } finally {
        if (!win.isDestroyed()) win.destroy()
      }
    },
  )

  ipcMain.handle('slides:recent', () => readRecent())

  // ── Show fullscreen: macOS native fullscreen is an animated Space transition, so
  // the slideshow would render windowed for ~1s mid-flight. Instead one call covers
  // everything while the show's black root hides the relayout: the tab view bleeds
  // over the tab strip (shell hook) and the window snaps via simpleFullScreen (same
  // trick as the audience window in presenter-show.ts). The renderer skips HTML
  // fullscreen on macOS entirely. Snap is skipped when the user already fullscreened
  // the window into its own Space; Windows/Linux keep HTML fullscreen (instant there)
  // and only need the bleed. Release is debounced: React strict-mode remounts and
  // presenter→show handoffs flip off→on within a tick, and honoring the off
  // immediately makes the window visibly bounce. ──
  let showFsRelease: ReturnType<typeof setTimeout> | null = null
  ipcMain.handle('slides:show-fullscreen', (e, on: boolean) => {
    const win = BrowserWindow.fromWebContents(e.sender) ?? windowRefs.shellWindow
    if (!win || win.isDestroyed()) return
    const wc = e.sender
    if (showFsRelease) {
      clearTimeout(showFsRelease)
      showFsRelease = null
    }
    if (on) {
      showChrome.setBleed?.(wc, true)
      if (process.platform === 'darwin' && !win.isFullScreen()) {
        win.setFullScreenable(false)
        if (!win.isSimpleFullScreen()) win.setSimpleFullScreen(true)
      }
      // The snap can leave the window's first responder on the shell chrome view, so
      // keys land in the tab-strip renderer (Esc dead until a click on the show).
      // win.focus() must NOT be used here — it focuses the shell renderer itself.
      // Focus the tab's webContents now and once more on the next tick (the snap's
      // responder change lands async). HTML fullscreen used to do this implicitly.
      wc.focus()
      setTimeout(() => {
        if (!wc.isDestroyed()) wc.focus()
      }, 50)
    } else {
      showFsRelease = setTimeout(() => {
        showFsRelease = null
        if (!wc.isDestroyed()) showChrome.setBleed?.(wc, false)
        if (win.isDestroyed()) return
        if (process.platform === 'darwin') {
          if (win.isSimpleFullScreen()) win.setSimpleFullScreen(false)
          win.setFullScreenable(true)
        }
      }, 150)
    }
  })

  // ── Chat attachments (slides:files-*) ──
  registerAttachmentIpc()

  // ── Presenter-view multi-screen show (registered inside registerSlidesIpc: shell
  // aggregate mode only calls this function) ──
  registerPresenterIpc()

  registerSlidesOnlyAiIpc()
}

// ── project-store IPC (standalone mode) ───────────────────────────────────
// In shell mode docs-main.registerProjectIpc registers these centrally (idempotent guard,
// registers once). Slides standalone calls this function.

let slidesProjectStore: ProjectStore | null = null
let slidesProjectIpcRegistered = false

function getSlidesProjectStore(): ProjectStore {
  if (!slidesProjectStore) slidesProjectStore = new ProjectStore(app.getPath('userData'))
  return slidesProjectStore
}

export function registerProjectIpc(): void {
  if (slidesProjectIpcRegistered) return
  slidesProjectIpcRegistered = true

  ipcMain.handle(
    'project:resolveChat',
    (_event, args: { filePath: string | null; tempChatId?: string }) => {
      const store = getSlidesProjectStore()
      store.ensureDefaultProject()
      if (!args.filePath) {
        return { projectId: 'default', chatId: args.tempChatId ?? `unsaved-${Date.now()}` }
      }
      return store.resolveChatForFile(args.filePath)
    },
  )

  ipcMain.handle(
    'project:appendChat',
    (
      _event,
      args: {
        projectId: string
        chatId: string
        role: 'user' | 'assistant'
        text: string
        tools?: Array<{
          name: string
          summary: string
          isError?: boolean
          input?: string
          output?: string
        }>
        attachments?: Array<{ name: string; path?: string; ext?: string; sizeBytes?: number }>
      },
    ) => {
      const msg: Parameters<ProjectStore['appendChatMessage']>[2] = {
        role: args.role,
        text: args.text,
      }
      if (args.tools) msg.tools = args.tools
      if (args.attachments) msg.attachments = args.attachments
      getSlidesProjectStore().appendChatMessage(args.projectId, args.chatId, msg)
    },
  )

  ipcMain.handle(
    'project:loadChat',
    (_event, args: { projectId: string; chatId: string; limit?: number }) => {
      return getSlidesProjectStore().loadChat(args.projectId, args.chatId, args.limit ?? 200)
    },
  )

  ipcMain.handle(
    'project:rebindChat',
    (
      _event,
      args: { projectId: string; tempChatId: string; newChatId?: string; newFilePath?: string },
    ) => {
      const store = getSlidesProjectStore()
      if (args.newFilePath) {
        return store.rebindChatToFile(args.projectId, args.tempChatId, args.newFilePath)
      }
      if (args.newChatId) store.rebindChat(args.projectId, args.tempChatId, args.newChatId)
      return { projectId: args.projectId, chatId: args.newChatId ?? args.tempChatId }
    },
  )
}

export function createSlidesWindow(openPath?: string | null): BrowserWindow {
  const win = new BrowserWindow({
    width: 1280,
    height: 840,
    title: 'GenOffice Slides',
    ...(process.platform === 'darwin'
      ? { titleBarStyle: 'hiddenInset' as const }
      : {
          titleBarStyle: 'hidden' as const,
          titleBarOverlay: { color: '#ffffff', symbolColor: '#444444', height: 40 },
        }),
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  trackSlidesWebContents(win.webContents)
  // Close guard for standalone-window mode (tab mode runs the same flow via the shell's tab-manager/window-close path)
  win.on('close', (event) => {
    if (!slidesIsDirty(win.webContents.id)) return
    event.preventDefault()
    void requestSlidesClose(win.webContents, win).then((proceed) => {
      // destroy() exits bypassing this handler (close() would re-enter the guard)
      if (proceed && !win.isDestroyed()) win.destroy()
    })
  })

  if (runtime.rendererDevUrl) win.loadURL(runtime.rendererDevUrl)
  else if (runtime.rendererFilePath) win.loadFile(runtime.rendererFilePath)

  if (openPath) {
    win.webContents.once('did-finish-load', async () => {
      try {
        const result = await openAndBuild(win.webContents, openPath, 1280)
        win.webContents.send('slides:opened', result)
      } catch {
        /* ignore */
      }
    })
  }
  return win
}

/** per-webContents background setter: opaque white while (re)loading, flipped
 * to transparent by consume-pending-open once the renderer has mounted, so the
 * vibrancy hole never shows the raw desktop behind an unpainted page */
const vibFlip = new Map<number, (color: string) => void>()

function armVibrancy(view: WebContentsView): void {
  if (process.platform !== 'darwin') return
  const setColor = (c: string) => view.setBackgroundColor(c)
  setColor('#ffffff')
  // view.webContents becomes undefined after destroy, so grab the id beforehand
  const wcId = view.webContents.id
  vibFlip.set(wcId, setColor)
  view.webContents.on('did-start-loading', () => setColor('#ffffff'))
  view.webContents.once('destroyed', () => vibFlip.delete(wcId))
}

/** Tab version of createSlidesWindow: same runtime/IPC, hosted in the shell's WebContentsView */
export function createSlidesView(openPath?: string | null): WebContentsView {
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  registerSlidesIpc()
  trackSlidesWebContents(view.webContents)
  armVibrancy(view)
  // The renderer calls consumePendingOpen on mount; use that to avoid a did-finish-load timing race
  if (openPath && existsSync(openPath)) pendingByWc.set(view.webContents.id, openPath)
  // mode=tab: the shell's tab strip owns the traffic lights / caption buttons,
  // so the ribbon must not reserve space for them
  if (runtime.rendererDevUrl) {
    // append via URL so a dev URL that already carries query params stays valid
    const devUrl = new URL(runtime.rendererDevUrl)
    devUrl.searchParams.set('mode', 'tab')
    void view.webContents.loadURL(devUrl.toString())
  } else if (runtime.rendererFilePath)
    void view.webContents.loadFile(runtime.rendererFilePath, { query: { mode: 'tab' } })
  return view
}

/** Items the shell injects into the File menu (e.g. Back to Home) */
let extraFileMenuItems: Electron.MenuItemConstructorOptions[] = []
export function setSlidesExtraFileMenuItems(items: Electron.MenuItemConstructorOptions[]): void {
  extraFileMenuItems = items
}

/** Tab mode: Cmd+W closes the current tab rather than the whole shell window */
let closeActiveTabHook: (() => void) | null = null
export function setSlidesCloseTabHook(fn: (() => void) | null): void {
  closeActiveTabHook = fn
}

export function buildSlidesMenu(): Menu {
  const send = (cmd: string) =>
    (windowRefs.activeWebContents ?? BrowserWindow.getFocusedWindow()?.webContents)?.send(
      'slides:menu',
      cmd,
    )
  const isMac = process.platform === 'darwin'
  const labels = appMenuLabels(getUiLang())
  const template: Electron.MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        { label: tm('menuOpen'), accelerator: 'CmdOrCtrl+O', click: () => send('open') },
        ...(extraFileMenuItems.length > 0
          ? [{ type: 'separator' as const }, ...extraFileMenuItems]
          : []),
        { type: 'separator' },
        { label: tm('menuSave'), accelerator: 'CmdOrCtrl+S', click: () => send('save') },
        { label: tm('menuSaveAs'), accelerator: 'CmdOrCtrl+Shift+S', click: () => send('save-as') },
        { type: 'separator' },
        // The ribbon's File tab is Windows-only, so without these macOS had no way
        // to export or print at all
        { label: tm('menuExportPdf'), click: () => send('export-pdf') },
        { label: tm('menuExportImages'), click: () => send('export-images') },
        { label: tm('menuPrint'), accelerator: 'CmdOrCtrl+P', click: () => send('print') },
        { type: 'separator' },
        closeActiveTabHook
          ? {
              label: isMac ? tm('menuClose') : tm('menuQuit'),
              accelerator: isMac ? 'CmdOrCtrl+W' : 'CmdOrCtrl+Q',
              click: () => closeActiveTabHook?.(),
            }
          : isMac
            ? { role: 'close' as const, label: tm('menuClose') }
            : { role: 'quit' as const, label: tm('menuQuit') },
      ],
    },
    {
      label: tm('menuEdit'),
      submenu: [
        // Undo/redo are sent to the renderer: text-editing state uses native execCommand, otherwise document history
        { label: tm('menuUndo'), accelerator: 'CmdOrCtrl+Z', click: () => send('undo') },
        { label: tm('menuRedo'), accelerator: 'Shift+CmdOrCtrl+Z', click: () => send('redo') },
        { type: 'separator' },
        // Cut/copy/paste forward the same way: in text state the renderer calls back to the native clipboard; in canvas state the element clipboard is used
        { label: tm('menuCut'), accelerator: 'CmdOrCtrl+X', click: () => send('cut') },
        { label: tm('menuCopy'), accelerator: 'CmdOrCtrl+C', click: () => send('copy') },
        { label: tm('menuPaste'), accelerator: 'CmdOrCtrl+V', click: () => send('paste') },
        { role: 'selectAll', label: labels.selectAll },
      ],
    },
    {
      label: tm('menuView'),
      submenu: [
        { label: tm('menuZoomIn'), accelerator: 'CmdOrCtrl+=', click: () => send('zoom-in') },
        { label: tm('menuZoomOut'), accelerator: 'CmdOrCtrl+-', click: () => send('zoom-out') },
        {
          label: tm('menuActualSize'),
          accelerator: 'CmdOrCtrl+0',
          click: () => send('zoom-reset'),
        },
        { type: 'separator' },
        toggleDevToolsItem(labels),
      ],
    },
  ]
  return Menu.buildFromTemplate(template)
}

export function installSlidesMenu(): void {
  Menu.setApplicationMenu(buildSlidesMenu())
}

/**
 * Attach a proxy to the main process's global fetch. Environment variables take priority;
 * otherwise, after app ready, read the system proxy via session.resolveProxy() (the critical
 * path for packaged builds launched by double-click).
 */
async function applyMainProcessProxy(): Promise<void> {
  const setDispatcher = async (proxyUrl: string) => {
    // spawned gsk CLI children do their own fetch and never see the
    // dispatcher below — forward the proxy to them via env
    setGskProxyUrl(proxyUrl)
    try {
      const { ProxyAgent, setGlobalDispatcher } = await import('undici')
      setGlobalDispatcher(new ProxyAgent(proxyUrl))
      // strip user:pass credentials before logging
      console.log('[proxy] main-process fetch via', proxyUrl.replace(/\/\/[^@/]*@/, '//***@'))
    } catch (e) {
      console.warn('[proxy] failed to set ProxyAgent:', e)
    }
  }
  const envProxy =
    process.env.HTTPS_PROXY ||
    process.env.https_proxy ||
    process.env.HTTP_PROXY ||
    process.env.http_proxy ||
    process.env.ALL_PROXY ||
    process.env.all_proxy
  if (envProxy) {
    await setDispatcher(envProxy)
    return
  }
  // No environment variables: read the system proxy (requires app ready)
  try {
    await app.whenReady()
    // PAC/rule proxies answer per-host: probe the host the login flow, the
    // Genspark LLM proxy and the gsk CLI actually target
    const resolved = await electronSession.defaultSession.resolveProxy('https://www.genspark.ai/')
    // resolveProxy returns strings like "PROXY 127.0.0.1:1087" or "DIRECT"
    const m = /PROXY\s+([^;]+)/i.exec(resolved || '')
    if (m) {
      await setDispatcher(`http://${m[1].trim()}`)
    } else {
      console.log('[proxy] system proxy = DIRECT, no dispatcher set')
    }
  } catch (e) {
    console.warn('[proxy] resolveProxy failed:', e)
  }
}

export function startSlidesStandalone(): void {
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  // Optional debug switch: enable CDP only in dev with SLIDES_CDP_PORT explicitly set (for
  // automated testing/troubleshooting); packaged builds (isPackaged) are unaffected.
  if (!app.isPackaged && process.env.SLIDES_CDP_PORT) {
    app.commandLine.appendSwitch('remote-debugging-port', process.env.SLIDES_CDP_PORT)
    app.commandLine.appendSwitch('remote-allow-origins', '*')
  }
  // GENOFFICE_USER_DATA: test drivers point this at a scratch dir so automated
  // instances get their own userData AND single-instance lock (the lock is scoped
  // to userData), allowing parallel instances alongside a normal dev run.
  if (!app.isPackaged && process.env.GENOFFICE_USER_DATA) {
    app.setPath('userData', process.env.GENOFFICE_USER_DATA)
  }
  // The main process's Node fetch (undici) does not use the system proxy by default, so access
  // from mainland China to overseas LLM APIs like api.anthropic.com hits ETIMEDOUT on direct
  // connections. Route the global dispatcher through the proxy; the renderer (Chromium) uses
  // the system proxy on its own and is unaffected. Prefer environment variables (terminal
  // launches); packaged builds launched by double-click don't inherit terminal environment
  // variables, so fall back to Electron session.resolveProxy() reading the system proxy
  // settings.
  void applyMainProcessProxy()
  if (!app.requestSingleInstanceLock()) {
    app.quit()
    return
  }

  app.on('open-file', (event, path) => {
    event.preventDefault()
    if (app.isReady()) {
      const win = BrowserWindow.getAllWindows()[0]
      if (win) {
        openAndBuild(win.webContents, path, 1280).then((r) =>
          win.webContents.send('slides:opened', r),
        )
        win.focus()
      } else createSlidesWindow(path)
    } else pendingOpenPath = path
  })

  const argPath = process.argv.find((a) => a.toLowerCase().endsWith('.pptx'))
  if (argPath && existsSync(argPath)) pendingOpenPath = argPath

  app.whenReady().then(async () => {
    setUiLang(normalizeLang(process.env.GENOFFICE_LANG ?? app.getLocale()))
    registerSlidesIpc()
    registerAiIpc()
    registerProjectIpc()
    Menu.setApplicationMenu(buildSlidesMenu())
    const win = createSlidesWindow(pendingOpenPath)
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createSlidesWindow()
    })

    // Test-only: with SLIDES_SMOKE_SHOT=/path, take a screenshot after loading and quit
    if (!app.isPackaged && process.env.SLIDES_SMOKE_SHOT) {
      win.webContents.once('did-finish-load', async () => {
        await new Promise((r) => setTimeout(r, 1800))
        try {
          const info = await win.webContents.executeJavaScript(
            `({ thumbs: document.querySelectorAll('.thumb').length,` +
              ` canvases: document.querySelectorAll('canvas').length,` +
              ` empty: !!document.querySelector('.empty') })`,
          )

          console.log('SMOKE_INFO=' + JSON.stringify(info))
          const png = await win.webContents.capturePage()
          const { writeFileSync } = await import('node:fs')
          writeFileSync(process.env.SLIDES_SMOKE_SHOT!, png.toPNG())

          console.log('SMOKE_SHOT_OK')
        } catch (e) {
          console.error('SMOKE_ERR', e)
        }
        app.quit()
      })
    }
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
}
