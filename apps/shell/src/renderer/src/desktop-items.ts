import type { DesktopPinInfo, FolderRoot, HomeApi, RecentEntry } from '../../shared/home-api'
import type { SettingsTarget } from './SettingsModal'
import type { DesktopStrings } from './desktop-strings'

// ── the unified item model ──────────────────────────────

/** the brand tiles an app item draws (files use FileBadge, folders a glass tile) */
export type AppTile =
  | 'docs'
  | 'sheets'
  | 'slides'
  | 'pdf'
  | 'markdown'
  | 'html'
  | 'ask'
  | 'templates'
  | 'plugins'
  | 'settings'

/**
 * One thing on the desktop, whichever source it came from — the model every
 * icon provider emits and the view renders. `tile` drives an app item's
 * glyph, `ext` a file item's FileBadge, and `path` the reveal/properties
 * affordances; the two flags carry the source's contract: `alias` items are
 * what the person dragged in (removable, OS-opened when the suite cannot),
 * `draggable` items move on the grid while widgets are placed by their own
 * window.
 */
export interface DesktopItem {
  readonly id: string
  readonly kind: 'app' | 'folder' | 'file' | 'widget'
  readonly title: string
  /** the brand tile an app item draws; null for every other kind */
  readonly tile: AppTile | null
  /** extension driving the FileBadge; '' for non-file kinds */
  readonly ext: string
  /** real path behind file / folder / alias items; null for apps and untitled widgets */
  readonly path: string | null
  /** last-modified time, for the modified-time sort and the properties dialog */
  readonly mtimeMs?: number
  /** a user-placed alias: removable from the desktop, opened by the OS when unknown */
  readonly alias: boolean
  /** the grid drag applies; widgets are placed by their window, not the grid */
  readonly draggable: boolean
  /** run the item: a tab, a parent surface, or the OS default */
  readonly open: () => void
}

// ── the provider registry ───────────────────────────────

/**
 * What a provider reads and where an item's actions route. Everything the
 * view owns flows in here, so the providers themselves stay pure: the same
 * registry serves any future host (the ask panel's desktop references, the
 * MCP `desktop` domain) without touching window state.
 */
export interface DesktopItemContext {
  readonly strings: DesktopStrings
  readonly roots: readonly FolderRoot[]
  readonly recents: readonly RecentEntry[]
  /** the dragged-in aliases, stat-resolved by the view */
  readonly aliases: readonly RecentEntry[]
  /** the alias paths, for cross-source dedupe */
  readonly aliasPaths: ReadonlySet<string>
  /** the document widgets pinned to the desktop */
  readonly pins: readonly DesktopPinInfo[]
  /** the bridge, partial-tolerant like the view's own calls */
  readonly api: Partial<HomeApi>
  readonly onOpenAsk: () => void
  readonly onOpenTemplates: () => void
  readonly onOpenSettings: (target: SettingsTarget) => void
  readonly onOpenFolder: (root: FolderRoot) => void
  /**
   * An alias the OS-open gate refused (the file vanished, the launch gate
   * said no): the view surfaces it as its transient surface notice so a
   * double-click that does nothing is at least said out loud. Optional so
   * other hosts (tests) need not supply it.
   */
  readonly onOpenExternalFailed?: () => void
}

/** one source of desktop items; the registry order is the display order */
export interface DesktopItemProvider {
  readonly id: 'apps' | 'folder-roots' | 'recents' | 'desktop-files' | 'widgets'
  getItems(ctx: DesktopItemContext): DesktopItem[]
}

/** the desktop's full source set, in display order */
export function defaultDesktopProviders(): DesktopItemProvider[] {
  return [
    createAppsProvider(),
    createFolderRootsProvider(),
    createRecentsProvider(),
    createDesktopFilesProvider(),
    createWidgetsProvider(),
  ]
}

/**
 * Fold every provider's items into the one list the view lays out, in
 * registry order. Cross-source dedupe is by path, first provider to claim a
 * path wins — the providers drop their own covered rows (a recent or folder
 * root that the person also dragged in as an alias), so an explicit alias
 * never fights a derived row for its slot.
 */
export function mergeDesktopItems(
  providers: readonly DesktopItemProvider[],
  ctx: DesktopItemContext,
): DesktopItem[] {
  const out: DesktopItem[] = []
  const seenPaths = new Set<string>()
  for (const provider of providers) {
    for (const item of provider.getItems(ctx)) {
      if (item.path !== null) {
        if (seenPaths.has(item.path)) continue
        seenPaths.add(item.path)
      }
      out.push(item)
    }
  }
  return out
}

// ── the providers ───────────────────────────────────────

function createAppsProvider(): DesktopItemProvider {
  return {
    id: 'apps',
    getItems(ctx) {
      const s = ctx.strings
      const api = ctx.api
      const apps: ReadonlyArray<{
        id: string
        tile: AppTile
        title: string
        open: () => void
      }> = [
        { id: 'app:docs', tile: 'docs', title: s.appDocs, open: () => void api.newDoc?.() },
        { id: 'app:sheets', tile: 'sheets', title: s.appSheets, open: () => void api.newSheet?.() },
        { id: 'app:slides', tile: 'slides', title: s.appSlides, open: () => void api.newSlide?.() },
        { id: 'app:pdf', tile: 'pdf', title: s.appPdf, open: () => void api.newPdf?.() },
        {
          id: 'app:markdown',
          tile: 'markdown',
          title: s.appMarkdown,
          open: () => void api.newMarkdown?.(),
        },
        { id: 'app:html', tile: 'html', title: s.appHtml, open: () => void api.newHtml?.() },
        { id: 'app:ask', tile: 'ask', title: s.appAsk, open: ctx.onOpenAsk },
        {
          id: 'app:templates',
          tile: 'templates',
          title: s.appTemplates,
          open: ctx.onOpenTemplates,
        },
        {
          id: 'app:plugins',
          tile: 'plugins',
          title: s.appPlugins,
          open: () => ctx.onOpenSettings({ section: 'plugins' }),
        },
        {
          id: 'app:settings',
          tile: 'settings',
          title: s.appSettings,
          open: () => ctx.onOpenSettings({ section: 'general' }),
        },
      ]
      return apps.map((app) => ({
        id: app.id,
        kind: 'app' as const,
        title: app.title,
        tile: app.tile,
        ext: '',
        path: null,
        alias: false,
        draggable: true,
        open: app.open,
      }))
    },
  }
}

function createFolderRootsProvider(): DesktopItemProvider {
  return {
    id: 'folder-roots',
    getItems(ctx) {
      return ctx.roots
        .filter((root) => !ctx.aliasPaths.has(root.path))
        .map((root) => ({
          id: `folder:${root.path}`,
          kind: 'folder' as const,
          title: root.name,
          tile: null,
          ext: '',
          path: root.path,
          alias: false,
          draggable: true,
          open: () => ctx.onOpenFolder(root),
        }))
    },
  }
}

function createRecentsProvider(): DesktopItemProvider {
  return {
    id: 'recents',
    getItems(ctx) {
      return ctx.recents
        .filter((entry) => !ctx.aliasPaths.has(entry.path))
        .map((entry) => ({
          id: `file:${entry.path}`,
          kind: 'file' as const,
          title: entry.name,
          tile: null,
          ext: entry.ext,
          path: entry.path,
          mtimeMs: entry.mtimeMs,
          alias: false,
          draggable: true,
          open: () => void ctx.api.openPath?.(entry.path),
        }))
    },
  }
}

/**
 * The real files and folders the person dragged onto the desktop — alias
 * semantics: the stored path is the whole item, nothing is copied or moved,
 * and removing the icon never touches the file on disk.
 *
 * A dropped folder opens in the OS file manager (the suite's own folder
 * browsing is the folder-roots mechanism, which stays untouched). Mapping the
 * OS home desktop is deliberately not special-cased: ~/Desktop joins like
 * any folder the person adds as a root — the in-app desktop is the agent's
 * work surface, not a Finder mirror (charter principle 3).
 */
function createDesktopFilesProvider(): DesktopItemProvider {
  return {
    id: 'desktop-files',
    getItems(ctx) {
      return ctx.aliases.map((entry) => ({
        id: `alias:${entry.path}`,
        kind: (entry.directory ? 'folder' : 'file') as 'folder' | 'file',
        title: entry.name,
        tile: null,
        ext: entry.ext,
        path: entry.path,
        mtimeMs: entry.mtimeMs,
        alias: true,
        draggable: true,
        open: () => {
          if (entry.directory || !isOpenableDocument(entry.path)) {
            // unknown to the suite (or a folder): the OS default app, on a
            // main-process path that only opens what the desktop holds; a
            // false answer is the gate refusing, which the view reports
            // rather than leaving the double-click a silent no-op
            const opened = ctx.api.desktopOpenExternal?.(entry.path)
            if (opened) {
              void opened.then(
                (ok) => {
                  if (!ok) ctx.onOpenExternalFailed?.()
                },
                () => undefined,
              )
            }
            return
          }
          void ctx.api.openPath?.(entry.path)
        },
      }))
    },
  }
}

/**
 * The document widgets pinned to the desktop. They are items of the model —
 * one per live pin — but not icons of the grid: the widget is its own live
 * window floating over the surface, so the view uses these entries to
 * reserve the cells underneath (widgetReservedCellKeys) and to know what a
 * pin's `open` would do; nothing is drawn for them.
 */
function createWidgetsProvider(): DesktopItemProvider {
  return {
    id: 'widgets',
    getItems(ctx) {
      return ctx.pins.map((pin) => ({
        id: pin.id,
        kind: 'widget' as const,
        title: pin.title,
        tile: null,
        ext: pin.kind,
        path: pin.filePath,
        alias: false,
        draggable: false,
        open: () => void ctx.api.desktopPinsActivate?.(pin.id),
      }))
    },
  }
}

// ── drop-in helpers ─────────────────────────────────────

/**
 * Extensions the suite opens in its own tabs. Renderer-side copy of
 * OPENABLE_DOC_RE in packages/electron-utils/src/drop-open.ts — the renderer
 * cannot import that module (it pulls `electron`), and the main-process
 * openExternal handler re-validates anyway, so this only picks the channel.
 */
const OPENABLE_DOC_RE =
  /\.(docx|odt|ott|xlsx|xlsm|xls|csv|tsv|ods|ots|pptx|odp|otp|pdf|md|markdown|html|htm)$/i

/** whether the suite opens this path in a tab rather than deferring to the OS */
export function isOpenableDocument(path: string): boolean {
  return OPENABLE_DOC_RE.test(path)
}

/** upper bound on how many aliases one drop may add (the drop bridge's own cap) */
const MAX_DROP_ALIASES = 20
/** one alias path is capped like every other path the bridge accepts */
const MAX_ALIAS_PATH_CHARS = 4096

/**
 * Resolve the files of an OS drop to real paths through the preload's
 * `webUtils.getPathForFile` bridge: a resolver failure (a sandboxed entry
 * Electron cannot map) skips that file, and duplicates of an already-resolved
 * path drop out — one alias per file.
 */
export function droppedFilePaths(
  files: ReadonlyArray<File>,
  pathForFile: (file: File) => string | undefined,
): string[] {
  const out: string[] = []
  for (const file of files) {
    if (out.length >= MAX_DROP_ALIASES) break
    let resolved: string
    try {
      resolved = (pathForFile(file) ?? '').trim()
    } catch {
      continue
    }
    if (!resolved || resolved.length > MAX_ALIAS_PATH_CHARS || out.includes(resolved)) continue
    out.push(resolved)
  }
  return out
}

// ── the widgets' avoidance band ─────────────────────────

/** grid cells one widget's reserve band covers (a 480×360 widget at the md pitch) */
export const WIDGET_RESERVE_CELLS = 12

/**
 * The grid's avoidance of pinned widgets, deliberately simple: widgets are
 * placed over the desktop's bottom-right corner,
 * so the icon grid reserves that corner — `WIDGET_RESERVE_CELLS` per widget,
 * filled from the bottom-right corner upward, one column at a time from the
 * right edge. The band's cells count as occupied for saved positions and the
 * free flow alike, so no icon ever flows under a widget; a widget dragged
 * elsewhere by hand may still cover icons, which is accepted at this level —
 * its window is on top by design.
 */
export function widgetReservedCellKeys(
  widgetCount: number,
  rows: number,
  cols: number,
): ReadonlySet<string> {
  const keys = new Set<string>()
  if (widgetCount <= 0 || rows < 1 || cols < 1) return keys
  const total = Math.min(widgetCount * WIDGET_RESERVE_CELLS, rows * cols)
  for (let i = 0; i < total; i += 1) {
    const col = cols - 1 - Math.floor(i / rows)
    if (col < 0) break
    const row = rows - 1 - (i % rows)
    keys.add(`${col}:${row}`)
  }
  return keys
}
