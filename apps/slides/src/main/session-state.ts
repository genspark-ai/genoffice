/**
 * Shared main-process state for GenOffice Slides, extracted from slides-main.ts so
 * the IPC modules (slides-main, ai-ipc, presenter-show) can share it:
 * per-renderer sessions, snapshot undo/redo history, runtime paths, window
 * references, and RenderSlide rebuild helpers.
 */
import { BrowserWindow, webContents } from 'electron'
import type { WebContents } from 'electron'
import { join } from 'node:path'
import { materializeSlide, type OpenedPptx, type Slide } from '@genoffice/pptx-engine'
import {
  buildRenderSlide,
  type FontMetricsProvider,
  type RenderSlide,
} from '@genoffice/pptx-render'
import { createSystemFontMetrics } from './fonts'
import { tiffToPng } from './tiff-decode'

export interface RuntimePaths {
  preloadPath: string
  rendererDevUrl?: string | undefined
  rendererFilePath?: string | undefined
}

export const runtime: RuntimePaths = {
  preloadPath: join(__dirname, '../preload/index.js'),
  rendererDevUrl: process.env.ELECTRON_RENDERER_URL,
  rendererFilePath: join(__dirname, '../renderer/index.html'),
}

export function configureSlidesRuntime(paths: RuntimePaths): void {
  runtime.preloadPath = paths.preloadPath
  runtime.rendererDevUrl = paths.rendererDevUrl
  runtime.rendererFilePath = paths.rendererFilePath
}

// One session per renderer process (standalone window or shell tab), keyed by webContents.id
export interface Session {
  path: string
  opened: OpenedPptx
  fitWidthPx: number
  undoStack: HistorySnapshot[]
  redoStack: HistorySnapshot[]
  /** Nested history transaction used to collapse an AI tool/run into one undo step. */
  historyBatch?: {
    depth: number
    undoStart: number
    before: HistorySnapshot
  }
  /** Rollback points for the AI panel's Snapshots list, keyed by id (one per AI run that edited the deck). */
  aiSnapshots?: Map<number, HistorySnapshot>
  /** Edits that only touch archive entries (notes/comments; element-level dirty cannot detect them), reset after save */
  metaDirty?: boolean
  /** Per-page PageVisualData from the HTML pipeline: appends rebuild "existing + new" wholesale.
      Set to null after any native edit — a rebuild would lose edits, so refuse to append instead. */
  htmlPages?: unknown[] | null
  /** Transform preview gesture in progress (the first preview already pushed an undo snapshot; later previews/final commit do not) */
  transformPreview?: boolean
  /** The part currently edited in master view (exception to the fidelity rule: only that part is written back) */
  masterEdit?: { partPath: string; slide: Slide } | null
  /** A history-state notification is already queued for this session (coalesces per task) */
  historyNotifyScheduled?: boolean
}
export const sessions = new Map<number, Session>()

// ── Undo/redo (snapshot-based) ─────────────────────────────────────────
// The document's source of truth lives in the main process (deck.slides mutated in place +
// archive.entries surgery), so history lives here too: snapshot both before every edit.
// slides needs a deep copy (elements are mutated in place); entries only needs a shallow Map
// copy (byte Buffers are never mutated in place, only replaced wholesale).
export interface HistorySnapshot {
  slides: Slide[]
  entries: Map<string, Uint8Array>
  size: { cx: number; cy: number }
}
const MAX_HISTORY = 50

function trimHistory(stack: HistorySnapshot[]): void {
  while (stack.length > MAX_HISTORY) stack.shift()
}

export function takeSnapshot(session: Session): HistorySnapshot {
  return {
    slides: structuredClone(session.opened.deck.slides),
    entries: new Map(session.opened.archive.entries),
    size: { ...session.opened.deck.size },
  }
}

// slides must be deep-copied: restoreSnapshot hands them to the live deck, which mutates in place
function cloneSnapshot(snap: HistorySnapshot): HistorySnapshot {
  return {
    slides: structuredClone(snap.slides),
    entries: new Map(snap.entries),
    size: { ...snap.size },
  }
}

/**
 * Tell the renderer whether undo/redo have anything to apply (drives the QAT
 * button gray states). Deferred with setImmediate so no-op handlers that push
 * a snapshot and then pop it back in the same turn report the settled state,
 * and multiple stack changes per turn coalesce into one message.
 */
export function scheduleHistoryNotify(session: Session): void {
  if (session.historyNotifyScheduled) return
  session.historyNotifyScheduled = true
  setImmediate(() => {
    session.historyNotifyScheduled = false
    for (const [id, s] of sessions) {
      if (s !== session) continue
      webContents.fromId(id)?.send('slides:history-changed', {
        canUndo: session.undoStack.length > 0,
        canRedo: session.redoStack.length > 0,
      })
      return
    }
  })
}

/** Call before an edit operation: push onto the undo stack and clear the redo stack. */
export function pushHistory(session: Session): void {
  session.undoStack.push(takeSnapshot(session))
  trimHistory(session.undoStack)
  session.redoStack = []
  session.htmlPages = null
  scheduleHistoryNotify(session)
}

/** Begin a nestable transaction. Individual edit handlers keep their normal rollback behavior. */
export function beginHistoryBatch(session: Session): void {
  if (session.historyBatch) {
    session.historyBatch.depth += 1
    return
  }
  session.historyBatch = {
    depth: 1,
    undoStart: session.undoStack.length,
    before: takeSnapshot(session),
  }
}

/**
 * End a transaction and collapse every successful edit since begin into the pre-transaction
 * snapshot. Failed/no-op handlers can continue popping their own snapshots safely.
 * Returns the pre-transaction snapshot when the outermost end collapsed real edits (null otherwise),
 * so the caller can register it as an AI-panel rollback point.
 */
export function endHistoryBatch(session: Session): HistorySnapshot | null {
  const batch = session.historyBatch
  if (!batch) return null
  batch.depth -= 1
  if (batch.depth > 0) return null
  session.historyBatch = undefined
  if (session.undoStack.length <= batch.undoStart) return null
  session.undoStack.splice(batch.undoStart)
  session.undoStack.push(batch.before)
  trimHistory(session.undoStack)
  scheduleHistoryNotify(session)
  return batch.before
}

/** Preserve the old deck and its history when AI replaces the entire presentation. */
export function carryHistoryForReplacement(
  previous: Session | undefined,
  replacement: Session,
): void {
  if (!previous) return
  pushHistory(previous)
  replacement.undoStack = previous.undoStack
  replacement.redoStack = previous.redoStack
  replacement.historyBatch = previous.historyBatch
  replacement.aiSnapshots = previous.aiSnapshots
  scheduleHistoryNotify(replacement)
}

const MAX_AI_SNAPSHOTS = 20
let nextAiSnapshotId = 1

/** Register a rollback point (stored as its own copy; `snap` typically also sits on the undo stack). */
export function registerAiSnapshot(session: Session, snap: HistorySnapshot): number {
  const map = (session.aiSnapshots ??= new Map())
  const id = nextAiSnapshotId++
  map.set(id, cloneSnapshot(snap))
  while (map.size > MAX_AI_SNAPSHOTS) map.delete(map.keys().next().value as number)
  return id
}

/** Roll the deck back to a registered AI snapshot; the pre-rollback state becomes one undo step. */
export function restoreAiSnapshot(session: Session, id: number): boolean {
  const snap = session.aiSnapshots?.get(id)
  if (!snap) return false
  pushHistory(session)
  restoreSnapshot(session, snap)
  session.aiSnapshots?.delete(id)
  return true
}

export function restoreSnapshot(session: Session, snap: HistorySnapshot): void {
  // Clone: the live deck mutates elements in place, so handing a snapshot's own
  // arrays over would let later edits rewrite history still referenced by the
  // other stack (undo → edit → redo would replay mutated state).
  const fresh = cloneSnapshot(snap)
  session.opened.deck.slides = fresh.slides
  session.opened.deck.size = fresh.size
  const entries = session.opened.archive.entries
  entries.clear()
  for (const [k, v] of fresh.entries) entries.set(k, v)
}

/**
 * Close a history batch that outlived its run: an AI tool path that
 * throws between begin and end would otherwise leave historyBatch set forever,
 * and undo/redo — which refuse to run mid-batch — would silently do nothing for
 * the rest of the session. Collapsing here keeps the run's edits as one step.
 */
export function settleStaleHistoryBatch(session: Session): void {
  while (session.historyBatch) {
    const collapsed = endHistoryBatch(session)
    if (collapsed) registerAiSnapshot(session, collapsed)
  }
}

// ── Window references (shell tab mode + active renderer tracking) ──────
export const windowRefs = {
  /** Parent window for dialogs in tab mode (the shell's single BrowserWindow) */
  shellWindow: null as BrowserWindow | null,
  /** Currently active slides renderer (window or tab view) — target of menu commands; the shell updates it on tab switch */
  activeWebContents: null as WebContents | null,
}

export function setSlidesShellWindow(win: BrowserWindow | null): void {
  windowRefs.shellWindow = win
}

export function setActiveSlidesWebContents(wc: WebContents | null): void {
  windowRefs.activeWebContents = wc
}

export function dialogParent(): BrowserWindow | undefined {
  return windowRefs.shellWindow ?? BrowserWindow.getFocusedWindow() ?? undefined
}

// ── RenderSlide rebuild helpers ─────────────────────────────────────────

/** Precise system-font metrics (lazily built, shared process-wide; unmatched fonts fall back to heuristics per run). */
let fontMetrics: FontMetricsProvider | null = null
export function getFontMetrics(): FontMetricsProvider {
  if (!fontMetrics) fontMetrics = createSystemFontMetrics()
  return fontMetrics
}

export function buildAllRenderSlides(opened: OpenedPptx, fitWidthPx: number): RenderSlide[] {
  return opened.deck.slides.map((s, i) =>
    buildRenderSlide(s, opened.deck.size, {
      fitWidthPx,
      media: makeMediaResolver(opened),
      metrics: getFontMetrics(),
      slideNo: i + 1,
    }),
  )
}

const DISPLAY_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
}

/** Image mediaRef -> dataUrl (lazily decoded). TIFF is transcoded to PNG for display
    (Chromium can't decode it); the archive keeps the original bytes for save fidelity. */
export function makeMediaResolver(opened: OpenedPptx) {
  const cache = new Map<string, string | undefined>()
  return (mediaRef: string): string | undefined => {
    if (cache.has(mediaRef)) return cache.get(mediaRef)
    const bytes = opened.archive.readBytes(mediaRef)
    let url: string | undefined
    if (bytes) {
      const ext = mediaRef.split('.').pop()?.toLowerCase() ?? 'png'
      if (ext === 'tif' || ext === 'tiff') {
        const decoded = tiffToPng(bytes)
        if (decoded) url = `data:image/png;base64,${Buffer.from(decoded.png).toString('base64')}`
      } else {
        const mime = DISPLAY_MIME[ext] ?? 'image/png'
        url = `data:${mime};base64,${Buffer.from(bytes).toString('base64')}`
      }
    }
    cache.set(mediaRef, url)
    return url
  }
}

/** Rebuild a single page's RenderSlide (sent back to the renderer after an edit). */
export function rebuildSlide(session: Session, slideIndex: number): RenderSlide | null {
  const slide = session.opened.deck.slides[slideIndex]
  if (!slide) return null
  return buildRenderSlide(slide, session.opened.deck.size, {
    fitWidthPx: session.fitWidthPx,
    media: makeMediaResolver(session.opened),
    metrics: getFontMetrics(),
    slideNo: slideIndex + 1,
  })
}

/**
 * Rebuild the RenderSlide after re-parsing the whole page (the model is stale after a chart
 * edit and needs a reparse). Equivalent to materializeSlide then rebuildSlide, but without the
 * structureDirty save logic.
 */
export function rebuildSlideWithReparse(session: Session, slideIndex: number): RenderSlide | null {
  const fresh = materializeSlide(session.opened, slideIndex)
  if (!fresh) return null
  return buildRenderSlide(fresh, session.opened.deck.size, {
    fitWidthPx: session.fitWidthPx,
    media: makeMediaResolver(session.opened),
    metrics: getFontMetrics(),
    slideNo: slideIndex + 1,
  })
}
