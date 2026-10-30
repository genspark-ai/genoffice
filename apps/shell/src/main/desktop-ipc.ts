import { existsSync, mkdirSync, readFileSync, readdirSync, realpathSync, statSync } from 'node:fs'
import { basename, extname, isAbsolute, join } from 'node:path'
import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron'
import type { IpcMainInvokeEvent, OpenDialogOptions } from 'electron'
import { atomicCopyFile } from './atomic-write'
import { createJsonStore } from './json-store'
import { readAppSettings, writeAppSetting } from './app-settings'
import { isProgramFile } from './tab-menu'
import { normalizeKbScope, type KbScope } from './knowledge-base'
import {
  HOME_CHANNELS,
  type DesktopLayout,
  type DesktopPinRecord,
  type DesktopPrefs,
} from '../shared/home-api'

/** where the desktop layout lives: one file the whole view reads and writes through */
export function desktopLayoutPath(userDataDir: string): string {
  return join(userDataDir, 'desktop-layout.json')
}

/** where the desktop's pinned widgets live, next to the layout */
export function desktopPinsPath(userDataDir: string): string {
  return join(userDataDir, 'desktop-pins.json')
}

/** where the desktop's look-and-feel prefs live, next to the layout */
export function desktopPrefsPath(userDataDir: string): string {
  return join(userDataDir, 'desktop-prefs.json')
}

/** the map never grows past this, whatever the renderer sends (defensive, not a refusal) */
const MAX_ENTRIES = 500
/** both coordinates are clamped into this range before anything is stored */
const MAX_COORD = 10000
/** how many aliases the desktop may hold, whatever the renderer sends */
const MAX_ALIASES = 100
/** one alias path is capped like every other path the bridge accepts */
const MAX_ALIAS_CHARS = 4096
/** the pin store's own bounds: a widget per document, a screen-sized rect at
 * most. Exported because the pin-creation gate in detached-windows.ts refuses
 * the (MAX_PINS + 1)-th widget here, before it can exist live but unrestorable. */
export const MAX_PINS = 8
const PIN_COORD_MAX = 10000
const PIN_SIZE_MIN = 240
const PIN_SIZE_MAX = 2400
const PIN_PATH_MAX = 4096

/** the prefs' own clamps: a wallpaper id can be longish (a custom file name),
 * the carousel cannot ask for more than one rotation a minute */
const MAX_WALLPAPER_ID = 200
const MAX_CAROUSEL_MINUTES = 24 * 60

const ICON_SIZES: ReadonlySet<string> = new Set(['sm', 'md', 'lg'])
const SORT_MODES: ReadonlySet<string> = new Set(['custom', 'name', 'kind', 'mtime'])
const WALLPAPER_FITS: ReadonlySet<string> = new Set([
  'cover',
  'contain',
  'stretch',
  'tile',
  'center',
])

/** the label color: '' (automatic) or a #rgb / #rrggbb hex */
const LABEL_COLOR_RE = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/

/** what a fresh profile (or a corrupt file) gets */
export const DEFAULT_DESKTOP_PREFS: DesktopPrefs = {
  wallpaper: 'default',
  carouselMinutes: 0,
  iconSize: 'md',
  sortMode: 'custom',
  wallpaperFit: 'cover',
  labelColor: '',
}

function isPosition(value: unknown): value is { x: number; y: number } {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const pos = value as Record<string, unknown>
  return (
    typeof pos.x === 'number' &&
    Number.isFinite(pos.x) &&
    typeof pos.y === 'number' &&
    Number.isFinite(pos.y)
  )
}

const clampCoord = (value: number): number => Math.min(MAX_COORD, Math.max(0, value))

/**
 * The store's invariant, applied to anything read or written: a plain object
 * keyed by icon id, each value exactly `{ x, y }` with both coordinates
 * clamped into range, at most `MAX_ENTRIES` of them, plus the alias path list
 * (non-empty absolute paths, deduped, at most `MAX_ALIASES`). A payload that
 * is not an object is the empty layout; entries that lost their shape are
 * dropped individually so one bad line never costs the rest of the map.
 */
function normalizeLayout(raw: unknown): DesktopLayout {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return emptyLayout()
  const source = (raw as { positions?: unknown }).positions
  if (source === null || typeof source !== 'object' || Array.isArray(source)) {
    return emptyLayout()
  }
  const positions: DesktopLayout['positions'] = {}
  for (const [id, value] of Object.entries(source as Record<string, unknown>)) {
    if (id === '' || !isPosition(value)) continue
    positions[id] = { x: clampCoord(value.x), y: clampCoord(value.y) }
    if (Object.keys(positions).length >= MAX_ENTRIES) break
  }
  return { positions, aliases: normalizeAliases((raw as { aliases?: unknown }).aliases) }
}

/** the layout an absent, corrupt or foreign file reads as */
function emptyLayout(): DesktopLayout {
  return { positions: {}, aliases: [] }
}

/**
 * The alias list's own invariant: non-empty absolute paths, deduped in
 * first-seen order, capped. Anything else a hand-edited file carries is
 * dropped entry by entry, like a position that lost its shape.
 */
function normalizeAliases(raw: unknown): string[] {
  if (!Array.isArray(raw)) return []
  const out: string[] = []
  for (const entry of raw) {
    if (typeof entry !== 'string') continue
    const path = entry.trim()
    if (!path || path.length > MAX_ALIAS_CHARS || !isAbsolute(path)) continue
    if (out.includes(path)) continue
    out.push(path)
    if (out.length >= MAX_ALIASES) break
  }
  return out
}

/**
 * A missing, corrupt or shapeless file is the empty layout — positions are a
 * convenience, never a reason Home fails to open. Reading is also the one
 * moment a hand-edited or foreign file is normalized into the invariant.
 * The store goes through the shared versioned json-store factory; version 1
 * is today's bare `{ positions, aliases }` body (see json-store.ts for the
 * marker/migration convention).
 */
const layoutStore = (userDataDir: string) =>
  createJsonStore<DesktopLayout>({
    path: () => desktopLayoutPath(userDataDir),
    version: 1,
    normalize: normalizeLayout,
  })

export function readDesktopLayout(userDataDir: string): DesktopLayout {
  return layoutStore(userDataDir).read()
}

// ── pinned widgets: the store detached-windows writes, restore reads ─────

/** the pin kinds the store accepts; anything else is dropped on read */
const PIN_KINDS: ReadonlySet<string> = new Set([
  'docs',
  'sheets',
  'slides',
  'pdf',
  'markdown',
  'html',
])

function isPinBounds(value: unknown): value is DesktopPinRecord['bounds'] {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const b = value as Record<string, unknown>
  return (
    typeof b.x === 'number' &&
    Number.isFinite(b.x) &&
    typeof b.y === 'number' &&
    Number.isFinite(b.y) &&
    typeof b.width === 'number' &&
    Number.isFinite(b.width) &&
    typeof b.height === 'number' &&
    Number.isFinite(b.height)
  )
}

/**
 * The pin store's invariant, applied to anything read: an array of at most
 * `MAX_PINS` records, each a known editor kind, a real-ish path and a finite
 * rect (coordinates clamped, sizes clamped into widget range). Duplicates by
 * file path keep the first — a pin's identity is its document.
 */
export function normalizePins(raw: unknown): DesktopPinRecord[] {
  if (!Array.isArray(raw)) return []
  const out: DesktopPinRecord[] = []
  for (const entry of raw) {
    if (entry === null || typeof entry !== 'object' || Array.isArray(entry)) continue
    const pin = entry as Record<string, unknown>
    if (typeof pin.kind !== 'string' || !PIN_KINDS.has(pin.kind)) continue
    if (typeof pin.filePath !== 'string' || !pin.filePath || !isAbsolute(pin.filePath)) continue
    if (pin.filePath.length > PIN_PATH_MAX) continue
    if (!isPinBounds(pin.bounds)) continue
    if (out.some((other) => other.filePath === pin.filePath)) continue
    out.push({
      kind: pin.kind as DesktopPinRecord['kind'],
      filePath: pin.filePath,
      bounds: {
        x: clampPinCoord(pin.bounds.x),
        y: clampPinCoord(pin.bounds.y),
        width: clampInt(pin.bounds.width, PIN_SIZE_MIN, PIN_SIZE_MAX),
        height: clampInt(pin.bounds.height, PIN_SIZE_MIN, PIN_SIZE_MAX),
      },
    })
    if (out.length >= MAX_PINS) break
  }
  return out
}

/** pin origins may sit on any display, so only magnitude is bounded */
const clampPinCoord = (value: number): number =>
  Math.min(PIN_COORD_MAX, Math.max(-PIN_COORD_MAX, value))

/**
 * A missing, corrupt or shapeless pins file is the empty list — a widget is a
 * convenience, never a reason the launch path fails. Version 1 is today's
 * `{ pins }` envelope: the file body wraps the list, the domain value is the
 * bare list (see json-store.ts for the convention).
 */
const pinsStore = (userDataDir: string) =>
  createJsonStore<DesktopPinRecord[]>({
    path: () => desktopPinsPath(userDataDir),
    version: 1,
    normalize: normalizePins,
    unwrap: (raw) => (raw as { pins?: unknown } | null)?.pins,
    wrap: (pins) => ({ pins }),
  })

export function readDesktopPins(userDataDir: string): DesktopPinRecord[] {
  return pinsStore(userDataDir).read()
}

/** the scratch-then-rename write every pin mutation funnels through */
export async function writeDesktopPins(
  userDataDir: string,
  pins: readonly DesktopPinRecord[],
): Promise<void> {
  await pinsStore(userDataDir).write([...pins])
}

/**
 * The prefs' read half: a missing, corrupt or shapeless file is the default
 * set — the desktop opens with its fallback look rather than failing.
 * Version 1 is today's bare prefs object (see json-store.ts).
 */
const prefsStore = (userDataDir: string) =>
  createJsonStore<DesktopPrefs>({
    path: () => desktopPrefsPath(userDataDir),
    version: 1,
    normalize: normalizePrefs,
  })

export function readDesktopPrefs(userDataDir: string): DesktopPrefs {
  return prefsStore(userDataDir).read()
}

/** the scratch-then-rename write both the IPC save and the CLI control path use */
export async function writeDesktopPrefs(userDataDir: string, prefs: DesktopPrefs): Promise<void> {
  await prefsStore(userDataDir).write(prefs)
}

/** the scratch-then-rename write the imports use, so a crash mid-write costs nothing */
async function writeLayout(userDataDir: string, layout: DesktopLayout): Promise<void> {
  await layoutStore(userDataDir).write(layout)
}

// ── the alias gate the open and add channels share ──────

/**
 * Whether the OS must never be handed this path: the tab menu's
 * program-extension list (the same blacklist `isProgramFile` applies),
 * extended over the two shapes a plain extension check misses. A program
 * package does not have to be a file — `.app` and `.workflow` bundles are
 * directories on macOS — so the check cannot be gated on `stat.isFile()`;
 * a directory with a program extension is exactly the dangerous shape and
 * is refused by its path alone. Symlinks are resolved with `realpathSync`
 * and the target's own extension re-checked, so a link named `notes.txt`
 * pointing at `Payload.app` is caught through what it actually launches;
 * `statSync` follows links, so the link's own name and its target are both
 * covered. A path whose real target cannot be resolved fails closed.
 */
export function isProgramLaunchPath(path: string): boolean {
  if (isProgramFile(path)) return true
  try {
    const real = realpathSync(path)
    // a non-symlink path resolves to itself (parent symlinks included), and
    // its basename survives the resolution — the re-check only ever adds
    // information for a link whose target has a different name
    return real !== path && isProgramFile(real)
  } catch {
    return true
  }
}

// ── wallpaper plumbing ─────────────────────────────────

/** one shipped wallpaper as wallpapers.json describes it */
export interface WallpaperEntry {
  id: string
  /** the renderer's string key for the localized display name */
  nameKey: string
  /** plain file name inside the wallpapers dir — the manifest never carries a path */
  filename: string
}

function isWallpaperEntry(value: unknown): value is WallpaperEntry {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const entry = value as Record<string, unknown>
  return (
    typeof entry.id === 'string' &&
    entry.id !== '' &&
    /^[a-z0-9-]+$/.test(entry.id) &&
    typeof entry.nameKey === 'string' &&
    entry.nameKey !== '' &&
    typeof entry.filename === 'string' &&
    // a plain PNG file name in the wallpapers dir, never a path
    entry.filename !== '' &&
    /\.png$/.test(entry.filename) &&
    !/[\\/]/.test(entry.filename) &&
    basename(entry.filename) === entry.filename
  )
}

/**
 * The shipped manifest, shape-checked, mirroring the templates manifest. A
 * missing or corrupt file is the empty list — a wallpaper is an offer, never a
 * reason the desktop fails to open — and entries that lost their shape are
 * dropped individually.
 */
export function readWallpaperManifest(dir: string): WallpaperEntry[] {
  let text: string
  try {
    text = readFileSync(join(dir, 'wallpapers.json'), 'utf8')
  } catch {
    return []
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    return []
  }
  if (!Array.isArray(parsed)) return []
  return parsed.filter(isWallpaperEntry)
}

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  webp: 'image/webp',
}

/** a wallpaper file as a data URL; an unreadable file answers '' */
function wallpaperDataUrl(path: string): string {
  let data: Buffer
  try {
    data = readFileSync(path)
  } catch {
    return ''
  }
  const mime = MIME_BY_EXT[extname(path).slice(1).toLowerCase()] ?? 'application/octet-stream'
  return `data:${mime};base64,${data.toString('base64')}`
}

/** the image extensions the picker, the custom dir and the CLI wallpaper path accept */
export const IMAGE_EXTENSIONS = ['png', 'jpg', 'jpeg', 'webp']

/**
 * The custom wallpaper's file name: `<timestamp>-<sanitized name>` — the
 * timestamp keeps two picks of the same image distinct, the sanitization (the
 * document tools' own) keeps the name a plain file name. The caller bumps the
 * suffix until the name is free.
 */
export function customWallpaperName(source: string, attempt: number): string {
  // basename() only splits the host platform's separator, so a Windows path
  // handed over on macOS/Linux (or a Windows build seeing a POSIX path) would
  // keep its directories baked into the name; normalize the separators first
  const local = source.replace(/\\/g, '/')
  const ext = extname(local).toLowerCase()
  const base = basename(local, ext)
    // eslint-disable-next-line no-control-regex -- control characters are rejected on purpose
    .replace(/[/\\:*?"<>|\u0000-\u001f]/g, '_')
    .trim()
    .slice(0, 80)
    .trim()
  const stem = attempt === 0 ? `${Date.now()}` : `${Date.now()}-${attempt + 1}`
  return `${stem}-${base || 'wallpaper'}${ext || '.png'}`
}

/** deps the tests inject; the defaults are the real Electron wiring */
export interface DesktopIpcDeps {
  /** dir holding the shipped wallpapers and wallpapers.json; tests inject a fixture dir */
  wallpapersDir?: () => string
  /** dir the picked custom wallpapers are copied into; default <userData>/wallpapers */
  customWallpapersDir?: () => string
  /** dir the layout and prefs files live in; default <userData> */
  userDataDir?: () => string
  /** the wallpaper image picker; default opens the OS dialog on the caller's window */
  pickImage?: (event: IpcMainInvokeEvent) => Promise<string | null>
}

/** dev ships the wallpapers from the app's own build tree; the packaged app
 * reads them from resources/wallpapers (electron-builder maps build/wallpapers there) */
export function defaultWallpapersDir(): string {
  return app.isPackaged
    ? join(process.resourcesPath, 'wallpapers')
    : join(app.getAppPath(), 'build', 'wallpapers')
}

/** the OS file picker over image files only, on the window that asked */
async function nativePickImage(event: IpcMainInvokeEvent): Promise<string | null> {
  const win = BrowserWindow.fromWebContents(event.sender)
  const opts: OpenDialogOptions = {
    title: 'Choose a wallpaper image',
    filters: [{ name: 'Images', extensions: IMAGE_EXTENSIONS }],
    properties: ['openFile'],
  }
  const result = await (win ? dialog.showOpenDialog(win, opts) : dialog.showOpenDialog(opts))
  return result.canceled ? null : (result.filePaths[0] ?? null)
}

/**
 * The Home desktop's IPC: the icon positions (read and whole-map save), the
 * dragged-in aliases (add and remove through their own validated channels;
 * the layout save cannot touch them), the alias OS-open (validated against
 * the stored alias list before the OS sees a path), and the desktop's
 * look-and-feel (prefs, wallpaper bytes, the custom-image pick).
 *
 * The renderer sends only ids and whole shapes, never paths: a wallpaper is
 * looked up in the shipped manifest or the custom dir by name, so a traversal
 * has no syntax to ride in on. The two path-shaped arguments are
 * `desktopOpenExternal` and `desktopAddAlias`, and both are gated before the
 * OS or the disk sees anything: an open is refused unless the path is already
 * in the stored alias list — the renderer cannot ask the OS to open anything
 * the person did not place on the desktop themselves — and an add is refused
 * unless the path exists, is absolute and is not a program package (the
 * launch gate both channels share). A compromised renderer therefore cannot
 * first write a path and then open it: the alias list is only ever grown by
 * the validated add channel, and a layout save that carries an `aliases`
 * field writes positions only, keeping the stored list (the field is ignored
 * rather than refused so a stale preload that still sends it keeps saving
 * positions fine). The layout and prefs files are the user's data: a save
 * that does not carry the promised shape is refused outright (the file stays
 * as it was), while size and coordinate range are only ever clamped — never
 * a user-facing refusal. Registered from index.ts like the other feature IPC
 * modules.
 *
 * One non-desktop pair rides along here because this module's register call
 * is already wired from index.ts: the Ask AI corpus scope (app-settings.json
 * `kbScope`) get/set channels, same raw-string kb: family the kb:list /
 * kb:search handlers use. They persist through the shared app-settings
 * helpers, so the scope the kb: readers pick up per call and the scope the
 * settings UI writes are one value on disk.
 */
export function registerDesktopIpc(deps: DesktopIpcDeps = {}): void {
  const userDataDir = deps.userDataDir ?? (() => app.getPath('userData'))
  const wallpapersDir = deps.wallpapersDir ?? defaultWallpapersDir
  const customWallpapersDir =
    deps.customWallpapersDir ?? (() => join(app.getPath('userData'), 'wallpapers'))
  const pickImage = deps.pickImage ?? nativePickImage

  // position saves and alias mutations are read-modify-write cycles over one
  // file; this chain keeps them from interleaving (a drop racing a debounced
  // position save must not cost either write)
  let layoutWrites: Promise<unknown> = Promise.resolve()
  const queueLayoutWrite = (write: () => Promise<void>): Promise<void> => {
    const next = layoutWrites.then(write, write)
    layoutWrites = next.catch(() => undefined)
    return next
  }

  ipcMain.handle(HOME_CHANNELS.desktopLayout, async (): Promise<DesktopLayout> =>
    readDesktopLayout(userDataDir()),
  )

  ipcMain.handle(HOME_CHANNELS.desktopLayoutSave, async (_event, raw: unknown): Promise<void> => {
    if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) {
      throw new Error('the desktop layout must be an object of { positions: { id: { x, y } } }')
    }
    const positions = (raw as { positions?: unknown }).positions
    if (positions === null || typeof positions !== 'object' || Array.isArray(positions)) {
      throw new Error('the desktop layout must be an object of { positions: { id: { x, y } } }')
    }
    for (const value of Object.values(positions as Record<string, unknown>)) {
      if (!isPosition(value)) {
        throw new Error('every desktop position must be { x: number, y: number }')
      }
    }
    await queueLayoutWrite(async () => {
      // the alias list is owned by the add/remove channels: whatever an
      // `aliases` field carries, the stored list survives this save untouched
      // (ignored, not refused, for the stale-preload reason above)
      const stored = readDesktopLayout(userDataDir())
      await writeLayout(userDataDir(), { ...normalizeLayout(raw), aliases: stored.aliases })
    })
  })

  // ── alias add/remove: the only channels that grow or shrink the list ──

  /**
   * The drop-in path's landing: the stored alias list grows by one validated
   * path. The whole check battery runs before the disk is touched — a string,
   * absolute, within the path cap, existing right now, and not a program
   * package (the launch gate the open channel shares, so nothing launchable
   * ever enters the whitelist). A refused path leaves the list as it was; the
   * answer is always the current list, so the renderer's next paint is the
   * truth (a refused drop simply never gains its icon). The one exception is
   * a list already at MAX_ALIASES: the drop is a person acting, not a corrupt
   * payload, so it is refused with an explicit error the UI can surface
   * instead of being silently dropped.
   */
  ipcMain.handle(
    HOME_CHANNELS.desktopAddAlias,
    async (_event, rawPath: unknown): Promise<readonly string[]> => {
      const current = (): readonly string[] => readDesktopLayout(userDataDir()).aliases
      if (typeof rawPath !== 'string' || !rawPath || rawPath.length > MAX_ALIAS_CHARS) {
        return current()
      }
      if (!isAbsolute(rawPath)) return current()
      try {
        statSync(rawPath)
      } catch {
        return current()
      }
      if (isProgramLaunchPath(rawPath)) return current()
      await queueLayoutWrite(async () => {
        const stored = readDesktopLayout(userDataDir())
        if (stored.aliases.includes(rawPath)) return
        if (stored.aliases.length >= MAX_ALIASES) {
          // the explicit refusal the renderer surfaces; a stale preload that
          // swallows rejections still answers the unchanged list, so the next
          // paint stays the truth either way
          throw new Error(
            `the desktop holds at most ${MAX_ALIASES} aliases; remove one before adding another`,
          )
        }
        const aliases = [...stored.aliases, rawPath]
        await writeLayout(userDataDir(), { ...stored, aliases })
      })
      return current()
    },
  )

  /**
   * The "Remove from desktop" action: the stored alias list drops one path. The file on
   * disk is never touched — the path is the whole item. An unknown path is a
   * no-op; the answer is the current list.
   */
  ipcMain.handle(
    HOME_CHANNELS.desktopRemoveAlias,
    async (_event, rawPath: unknown): Promise<readonly string[]> => {
      const current = (): readonly string[] => readDesktopLayout(userDataDir()).aliases
      if (typeof rawPath !== 'string' || !rawPath || rawPath.length > MAX_ALIAS_CHARS) {
        return current()
      }
      if (!isAbsolute(rawPath)) return current()
      if (!current().includes(rawPath)) return current()
      await queueLayoutWrite(async () => {
        const stored = readDesktopLayout(userDataDir())
        const aliases = stored.aliases.filter((candidate) => candidate !== rawPath)
        await writeLayout(userDataDir(), { ...stored, aliases })
      })
      return current()
    },
  )

  // ── alias opens: the OS default app, only for what the person placed ──

  /**
   * The alias double-click's unknown-type path: the file the OS should open.
   * Everything is validated before the OS sees it — the path must be one the
   * desktop actually holds (a renderer cannot ask for arbitrary launches),
   * it must exist right now, and a program package never launches through
   * here (the launch gate: the same extension list the tab menu's "Open with
   * the default app" uses, plus directories and resolved symlinks).
   */
  ipcMain.handle(
    HOME_CHANNELS.desktopOpenExternal,
    async (_event, rawPath: unknown): Promise<boolean> => {
      if (typeof rawPath !== 'string' || !rawPath || rawPath.length > MAX_ALIAS_CHARS) return false
      if (!isAbsolute(rawPath)) return false
      if (!readDesktopLayout(userDataDir()).aliases.includes(rawPath)) return false
      try {
        statSync(rawPath)
      } catch {
        return false
      }
      if (isProgramLaunchPath(rawPath)) return false
      // openPath answers '' on success; a failure string is logged nowhere else,
      // so it is dropped here the way the reveal path drops it
      await shell.openPath(rawPath)
      return true
    },
  )

  // ── prefs ──

  ipcMain.handle(HOME_CHANNELS.desktopPrefs, async (): Promise<DesktopPrefs> =>
    readDesktopPrefs(userDataDir()),
  )

  ipcMain.handle(HOME_CHANNELS.desktopPrefsSave, async (_event, raw: unknown): Promise<void> => {
    if (!isDesktopPrefs(raw)) {
      throw new Error(
        'the desktop prefs must be { wallpaper: string, carouselMinutes: number, ' +
          "iconSize: 'sm'|'md'|'lg', sortMode: 'custom'|'name'|'kind'|'mtime', " +
          "wallpaperFit: 'cover'|'contain'|'stretch'|'tile'|'center', labelColor: string }",
      )
    }
    await writeDesktopPrefs(userDataDir(), normalizePrefs(raw))
  })

  // ── wallpapers ──

  ipcMain.handle(HOME_CHANNELS.wallpaperImage, async (_event, rawId: unknown): Promise<string> => {
    if (typeof rawId !== 'string') return ''
    // 'default' is the live CSS aurora: there is no file behind it
    if (rawId === '' || rawId === 'default') return ''
    if (rawId.startsWith('builtin:')) {
      // the file name comes from the shipped manifest, never from the id
      const entry = readWallpaperManifest(wallpapersDir()).find(
        (candidate) => candidate.id === rawId.slice('builtin:'.length),
      )
      if (!entry) return ''
      return wallpaperDataUrl(join(wallpapersDir(), entry.filename))
    }
    if (rawId.startsWith('custom:')) {
      // a plain file name that must already sit in the custom dir; anything
      // path-like (../, ..\, a separator) never matches a readdir entry
      const name = rawId.slice('custom:'.length)
      if (name === '' || name !== basename(name) || /[\\/]/.test(name)) return ''
      let present: string[]
      try {
        present = readdirSync(customWallpapersDir())
      } catch {
        return ''
      }
      if (!present.includes(name)) return ''
      return wallpaperDataUrl(join(customWallpapersDir(), name))
    }
    return ''
  })

  ipcMain.handle(HOME_CHANNELS.wallpaperPickCustom, async (event): Promise<string> => {
    const source = await pickImage(event)
    if (!source) return ''
    const dir = customWallpapersDir()
    mkdirSync(dir, { recursive: true })
    // the timestamped name is tried with a bumping suffix until it is free,
    // so re-picking the same image in the same instant never clobbers it
    let attempt = 0
    let name = customWallpaperName(source, attempt)
    while (existsSync(join(dir, name))) {
      attempt += 1
      name = customWallpaperName(source, attempt)
    }
    await atomicCopyFile(source, join(dir, name))
    return `custom:${name}`
  })

  // ── the Ask AI corpus scope: app-settings.json `kbScope` ──

  /** the settings file this module's scope pair reads and writes */
  const appSettingsPath = (): string => join(userDataDir(), 'app-settings.json')

  /** the persisted corpus scope, normalized exactly like the kb: readers do */
  ipcMain.handle('kb:get-scope', async (): Promise<KbScope> =>
    normalizeKbScope(readAppSettings(appSettingsPath()).kbScope),
  )

  /** set the corpus scope; anything but the two known values is refused so a
   * junk write can never land in the file the kb: readers trust */
  ipcMain.handle('kb:set-scope', async (_event, raw: unknown): Promise<KbScope> => {
    if (raw !== 'indexed' && raw !== 'starred') {
      throw new Error("the corpus scope must be 'indexed' or 'starred'")
    }
    writeAppSetting(appSettingsPath(), 'kbScope', raw)
    return raw
  })
}

/**
 * The prefs' invariant, applied to anything read or written: exactly the six
 * contract fields, every one valid — an unknown enum value, a negative carousel
 * or a non-hex label color falls back to its default, like the layout's
 * per-entry clamp. A payload that is not an object at all is the full default set.
 * Exported for the CLI's desktop domain, which writes through the same
 * invariant on its set-wallpaper path.
 */
export function normalizePrefs(raw: unknown): DesktopPrefs {
  const prefs = { ...DEFAULT_DESKTOP_PREFS }
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return prefs
  const source = raw as Record<string, unknown>
  if (
    typeof source.wallpaper === 'string' &&
    source.wallpaper !== '' &&
    source.wallpaper.length <= MAX_WALLPAPER_ID
  ) {
    prefs.wallpaper = source.wallpaper
  }
  if (typeof source.carouselMinutes === 'number' && Number.isFinite(source.carouselMinutes)) {
    prefs.carouselMinutes = clampInt(Math.round(source.carouselMinutes), 0, MAX_CAROUSEL_MINUTES)
  }
  if (typeof source.iconSize === 'string' && ICON_SIZES.has(source.iconSize)) {
    prefs.iconSize = source.iconSize as DesktopPrefs['iconSize']
  }
  if (typeof source.sortMode === 'string' && SORT_MODES.has(source.sortMode)) {
    prefs.sortMode = source.sortMode as DesktopPrefs['sortMode']
  }
  if (typeof source.wallpaperFit === 'string' && WALLPAPER_FITS.has(source.wallpaperFit)) {
    prefs.wallpaperFit = source.wallpaperFit as DesktopPrefs['wallpaperFit']
  }
  if (
    typeof source.labelColor === 'string' &&
    (source.labelColor === '' || LABEL_COLOR_RE.test(source.labelColor))
  ) {
    prefs.labelColor = source.labelColor
  }
  return prefs
}

/**
 * The save's refusal check: all six fields present and of the promised shape.
 * Extra keys are dropped by the write (the layout's policy); missing or
 * malformed ones are refused so the file can never hold half a prefs.
 */
function isDesktopPrefs(value: unknown): value is DesktopPrefs {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const prefs = value as Record<string, unknown>
  return (
    typeof prefs.wallpaper === 'string' &&
    prefs.wallpaper !== '' &&
    prefs.wallpaper.length <= MAX_WALLPAPER_ID &&
    typeof prefs.carouselMinutes === 'number' &&
    Number.isFinite(prefs.carouselMinutes) &&
    prefs.carouselMinutes >= 0 &&
    typeof prefs.iconSize === 'string' &&
    ICON_SIZES.has(prefs.iconSize) &&
    typeof prefs.sortMode === 'string' &&
    SORT_MODES.has(prefs.sortMode) &&
    typeof prefs.wallpaperFit === 'string' &&
    WALLPAPER_FITS.has(prefs.wallpaperFit) &&
    typeof prefs.labelColor === 'string' &&
    (prefs.labelColor === '' || LABEL_COLOR_RE.test(prefs.labelColor))
  )
}

function clampInt(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max)
}
