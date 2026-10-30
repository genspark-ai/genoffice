import { existsSync, mkdirSync, readFileSync, readdirSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import { app, BrowserWindow, dialog, ipcMain } from 'electron'
import type { IpcMainInvokeEvent, OpenDialogOptions } from 'electron'
import { atomicCopyFile, atomicWriteFile } from './atomic-write'
import { HOME_CHANNELS, type DesktopLayout, type DesktopPrefs } from '../shared/home-api'

/** where the desktop layout lives: one file the whole view reads and writes through */
export function desktopLayoutPath(userDataDir: string): string {
  return join(userDataDir, 'desktop-layout.json')
}

/** where the desktop's look-and-feel prefs live, next to the layout */
export function desktopPrefsPath(userDataDir: string): string {
  return join(userDataDir, 'desktop-prefs.json')
}

/** the map never grows past this, whatever the renderer sends (defensive, not a refusal) */
const MAX_ENTRIES = 500
/** both coordinates are clamped into this range before anything is stored */
const MAX_COORD = 10000

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
 * clamped into range, at most `MAX_ENTRIES` of them. A payload that is not an
 * object is the empty layout; entries that lost their shape are dropped
 * individually so one bad line never costs the rest of the map.
 */
function normalizeLayout(raw: unknown): DesktopLayout {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return { positions: {} }
  const source = (raw as { positions?: unknown }).positions
  if (source === null || typeof source !== 'object' || Array.isArray(source)) {
    return { positions: {} }
  }
  const positions: DesktopLayout['positions'] = {}
  for (const [id, value] of Object.entries(source as Record<string, unknown>)) {
    if (id === '' || !isPosition(value)) continue
    positions[id] = { x: clampCoord(value.x), y: clampCoord(value.y) }
    if (Object.keys(positions).length >= MAX_ENTRIES) break
  }
  return { positions }
}

/**
 * A missing, corrupt or shapeless file is the empty layout — positions are a
 * convenience, never a reason Home fails to open. Reading is also the one
 * moment a hand-edited or foreign file is normalized into the invariant.
 */
export function readDesktopLayout(userDataDir: string): DesktopLayout {
  let text: string
  try {
    text = readFileSync(desktopLayoutPath(userDataDir), 'utf8')
  } catch {
    return { positions: {} }
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    return { positions: {} }
  }
  return normalizeLayout(parsed)
}

/**
 * The prefs' read half: a missing, corrupt or shapeless file is the default
 * set — the desktop opens with its fallback look rather than failing.
 */
export function readDesktopPrefs(userDataDir: string): DesktopPrefs {
  let text: string
  try {
    text = readFileSync(desktopPrefsPath(userDataDir), 'utf8')
  } catch {
    return { ...DEFAULT_DESKTOP_PREFS }
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    return { ...DEFAULT_DESKTOP_PREFS }
  }
  return normalizePrefs(parsed)
}

/** the scratch-then-rename write both the IPC save and the CLI control path use */
export async function writeDesktopPrefs(userDataDir: string, prefs: DesktopPrefs): Promise<void> {
  await atomicWriteFile(desktopPrefsPath(userDataDir), Buffer.from(JSON.stringify(prefs, null, 2)))
}

/** the scratch-then-rename write the imports use, so a crash mid-write costs nothing */
async function writeLayout(userDataDir: string, layout: DesktopLayout): Promise<void> {
  await atomicWriteFile(
    desktopLayoutPath(userDataDir),
    Buffer.from(JSON.stringify(layout, null, 2)),
  )
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
 * The Home desktop's IPC: the icon positions (read and whole-map save) and the
 * desktop's look-and-feel (prefs, wallpaper bytes, the custom-image pick).
 *
 * The renderer sends only ids and whole shapes, never paths: a wallpaper is
 * looked up in the shipped manifest or the custom dir by name, so a traversal
 * has no syntax to ride in on. The layout and prefs files are the user's data:
 * a save that does not carry the promised shape is refused outright (the file
 * stays as it was), while size and coordinate range are only ever clamped —
 * never a user-facing refusal. Registered from index.ts like the other feature
 * IPC modules.
 */
export function registerDesktopIpc(deps: DesktopIpcDeps = {}): void {
  const userDataDir = deps.userDataDir ?? (() => app.getPath('userData'))
  const wallpapersDir = deps.wallpapersDir ?? defaultWallpapersDir
  const customWallpapersDir =
    deps.customWallpapersDir ?? (() => join(app.getPath('userData'), 'wallpapers'))
  const pickImage = deps.pickImage ?? nativePickImage

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
    await writeLayout(userDataDir(), normalizeLayout(raw))
  })

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
