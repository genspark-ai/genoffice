import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type {
  CSSProperties,
  PointerEvent as ReactPointerEvent,
  ReactElement,
  ReactNode,
} from 'react'
import { createPortal } from 'react-dom'
import type { Lang } from '@genoffice/i18n'
import { useDismissablePopover } from '@genoffice/ui'
import type { DesktopPrefs, FolderRoot, HomeApi, RecentEntry } from '../../shared/home-api'
import { DesktopPropsDialog, IconPropsDialog } from './DesktopProperties'
import type { DesktopPropsTarget } from './DesktopProperties'
import { ANIMATED_WALLPAPERS, DesktopWallpaperCenter } from './DesktopWallpaperCenter'
import { FileBadge } from './file-badge'
import { placeFolderMenu } from './folder-menu-position'
import type { SettingsTarget } from './SettingsModal'
import { desktopStrings } from './desktop-strings'
import type { DesktopStrings } from './desktop-strings'
import './desktop.css'

// ── grid geometry ───────────────────────────────────────
// One icon's footprint: a 64px glyph tile plus a two-line label underneath.
// Every position the layout file stores is the top-left pixel of such a cell,
// which is what makes saved positions, grid snap and free-cell flow all agree.

/** horizontal cell pitch of the icon grid at the medium icon size */
export const DESKTOP_CELL_W = 96
/** vertical cell pitch of the icon grid at the medium icon size */
export const DESKTOP_CELL_H = 104
/** inset of the first cell from the surface's top-left */
export const DESKTOP_PAD_X = 12
export const DESKTOP_PAD_Y = 12
/** write-through window: a burst of rearrangements costs one save */
const SAVE_DEBOUNCE_MS = 300
/** recents shown as icons, newest first */
const RECENT_LIMIT = 8
/** pointer travel before a press becomes a drag */
const DRAG_THRESHOLD = 4
/** rows assumed before the surface is measured (one paint, then the real count) */
const UNMEASURED_ROWS = 5
/** pointer id a synthetic (test) mouse event carries */
const MOUSE_POINTER_ID = 1
/** a double-click this soon after a drop is a stray, not an open */
const DROP_DBLCLICK_GUARD_MS = 400

/** the icon-size multipliers: the cell pitch and the CSS --desktop-icon-scale
 * (desktop.css's icons-sm/md/lg) read the same numbers, so a cell always fits
 * its contents and saved positions re-quantize onto the same rows and columns */
const ICON_SCALES: Record<DesktopPrefs['iconSize'], number> = { sm: 0.75, md: 1, lg: 1.25 }

/** the 新建 submenu: the six document types and their new* bridge calls */
const NEW_DOCUMENT_ACTIONS = [
  { action: 'newDoc', labelKey: 'appDocs' },
  { action: 'newSheet', labelKey: 'appSheets' },
  { action: 'newSlide', labelKey: 'appSlides' },
  { action: 'newPdf', labelKey: 'appPdf' },
  { action: 'newMarkdown', labelKey: 'appMarkdown' },
  { action: 'newHtml', labelKey: 'appHtml' },
] as const satisfies ReadonlyArray<{
  action: 'newDoc' | 'newSheet' | 'newSlide' | 'newPdf' | 'newMarkdown' | 'newHtml'
  labelKey: keyof DesktopStrings
}>

/** the 排序方式 submenu: the read-only flows plus the free-drag default */
const SORT_MODE_ITEMS = [
  { mode: 'custom', labelKey: 'sortCustom' },
  { mode: 'name', labelKey: 'sortName' },
  { mode: 'kind', labelKey: 'sortKind' },
  { mode: 'mtime', labelKey: 'sortMtime' },
] as const satisfies ReadonlyArray<{
  mode: DesktopPrefs['sortMode']
  labelKey: keyof DesktopStrings
}>

/** the 查看 ▸ 图标大小 submenu */
const ICON_SIZE_ITEMS = [
  { size: 'sm', labelKey: 'sizeSm' },
  { size: 'md', labelKey: 'sizeMd' },
  { size: 'lg', labelKey: 'sizeLg' },
] as const satisfies ReadonlyArray<{
  size: DesktopPrefs['iconSize']
  labelKey: keyof DesktopStrings
}>

/** one fit's background-* mapping, applied inline to every wallpaper layer so
 * the crossfading pair (and the whole mode) always agrees */
const WALLPAPER_FIT_STYLE: Record<DesktopPrefs['wallpaperFit'], CSSProperties> = {
  cover: { backgroundSize: 'cover', backgroundPosition: 'center', backgroundRepeat: 'no-repeat' },
  contain: {
    backgroundSize: 'contain',
    backgroundPosition: 'center',
    backgroundRepeat: 'no-repeat',
  },
  stretch: {
    backgroundSize: '100% 100%',
    backgroundPosition: 'center',
    backgroundRepeat: 'no-repeat',
  },
  tile: { backgroundSize: 'auto', backgroundPosition: 'left top', backgroundRepeat: 'repeat' },
  center: { backgroundSize: 'auto', backgroundPosition: 'center', backgroundRepeat: 'no-repeat' },
}

/** the static built-ins the timed carousel crossfades through */
const STATIC_WALLPAPERS = [
  'builtin:brand-aurora',
  'builtin:warm-dusk',
  'builtin:dot-grid-dark',
  'builtin:minimal-light',
  'builtin:teal-deep',
] as const

/** what a fresh profile (or a preload that predates the prefs) gets */
const DEFAULT_PREFS: DesktopPrefs = {
  wallpaper: 'default',
  carouselMinutes: 0,
  iconSize: 'md',
  sortMode: 'custom',
  wallpaperFit: 'cover',
  labelColor: '',
}

type AppTile =
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

interface DesktopIcon {
  readonly id: string
  readonly kind: 'app' | 'file' | 'folder'
  readonly label: string
  /** the brand tile an app icon draws; files use FileBadge, folders a glass tile */
  readonly tile: AppTile | null
  /** extension driving the FileBadge of a file icon */
  readonly ext: string
  /** real path of a file / folder icon (context-menu reveal, the folder tooltip) */
  readonly path: string | null
  /** last-modified time of a file icon, for the 修改时间 sort and the 属性 dialog */
  readonly mtimeMs?: number
  readonly run: () => void
}

interface IconPos {
  x: number
  y: number
}

export interface DesktopViewProps {
  /** the UI language, from the locale context the parent holds */
  readonly lang: Lang
  /** the Home folder roots, already loaded by the parent */
  readonly roots: readonly FolderRoot[]
  /** switch Home to the AI ask view */
  readonly onOpenAsk: () => void
  /** open the template gallery overlay */
  readonly onOpenTemplates: () => void
  /** open Home's settings modal on a section */
  readonly onOpenSettings: (target: SettingsTarget) => void
  /** open one folder root's list view */
  readonly onOpenFolder: (root: FolderRoot) => void
  /** the Home-root node the wallpaper backdrop portals into, so it paints
   * behind the sidebar too; undefined (standalone) renders it in place */
  readonly backdropHost?: HTMLDivElement | null
}

const pointerIdOf = (event: PointerEvent | ReactPointerEvent): number =>
  typeof event.pointerId === 'number' ? event.pointerId : MOUSE_POINTER_ID

const clampInt = (value: number, min: number, max: number): number =>
  Math.min(Math.max(value, min), Math.max(min, max))

/** case-insensitive codepoint order: deterministic across machines and ICU builds */
const compareText = (a: string, b: string): number => {
  const left = a.toLowerCase()
  const right = b.toLowerCase()
  return left < right ? -1 : left > right ? 1 : 0
}

/** app icons first (system furniture, never "modified"), then folders, then files */
const kindRank = (icon: DesktopIcon): number =>
  icon.kind === 'app' ? 0 : icon.kind === 'folder' ? 1 : 2

/**
 * The read-only sort flows (名称 / 类型 / 修改时间): one stable ordering of the
 * icon set, laid out by the same free-cell flow as the default view. 'custom'
 * is the free-drag layout and returns the icon set untouched. Ties always fall
 * back to the icon id, so two sorts of one set are one order.
 */
function sortDesktopIcons(
  icons: readonly DesktopIcon[],
  mode: DesktopPrefs['sortMode'],
): DesktopIcon[] {
  if (mode === 'custom') return [...icons]
  const ranked = [...icons]
  if (mode === 'name') {
    ranked.sort((a, b) => compareText(a.label, b.label) || compareText(a.id, b.id))
  } else if (mode === 'kind') {
    ranked.sort(
      (a, b) =>
        kindRank(a) - kindRank(b) ||
        compareText(a.ext, b.ext) ||
        compareText(a.label, b.label) ||
        compareText(a.id, b.id),
    )
  } else {
    // 修改时间: system icons have no mtime and rank as never-modified (first,
    // in their built order), files newest first
    const mtimeOf = (icon: DesktopIcon): number => icon.mtimeMs ?? Number.POSITIVE_INFINITY
    ranked.sort((a, b) => mtimeOf(b) - mtimeOf(a) || compareText(a.id, b.id))
  }
  return ranked
}

/** prefers-reduced-motion, guarded for test DOMs that lack matchMedia */
function prefersReducedMotion(): boolean {
  return (
    typeof window !== 'undefined' &&
    typeof window.matchMedia === 'function' &&
    window.matchMedia('(prefers-reduced-motion: reduce)').matches
  )
}

/**
 * Relative luminance (WCAG) of a #rgb / #rrggbb hex, 0 (black) to 1 (white).
 * A custom label color above the threshold reads as light ink, so its scrim
 * darkens (a shadow); below it the scrim lightens (a white glow).
 */
function hexLuminance(hex: string): number {
  const digits =
    hex.length === 4
      ? [hex[1]! + hex[1]!, hex[2]! + hex[2]!, hex[3]! + hex[3]!]
      : [hex.slice(1, 3), hex.slice(3, 5), hex.slice(5, 7)]
  const channel = (value: string): number => {
    const raw = Number.parseInt(value, 16) / 255
    return raw <= 0.03928 ? raw / 12.92 : ((raw + 0.055) / 1.055) ** 2.4
  }
  return 0.2126 * channel(digits[0]!) + 0.7152 * channel(digits[1]!) + 0.0722 * channel(digits[2]!)
}

/** the label-color scrim that keeps a custom ink readable: light ink shades
 * into the dark scrim, dark ink into the light one (desktop.css defines both) */
const labelScrimOf = (labelColor: string): string =>
  hexLuminance(labelColor) > 0.5
    ? 'var(--desktop-label-scrim-light-ink)'
    : 'var(--desktop-label-scrim-dark-ink)'

/** builtins whose artwork is dark enough that automatic label ink goes light */
const DARK_WALLPAPERS: ReadonlySet<string> = new Set([
  'builtin:brand-aurora',
  'builtin:starfield-night',
  'builtin:warm-dusk',
  'builtin:dot-grid-dark',
  'builtin:teal-deep',
])

/**
 * Saved positions are cell origins — but a hand-edited or foreign file can
 * carry anything, so every position read back is quantized onto the grid
 * before it is used. That keeps the no-overlap invariant: an icon can only
 * ever sit at a cell origin, so two icons can only collide by claiming the
 * same cell, and the layout algorithm never lets that happen. The pitch is
 * the current icon size's, so positions survive a size switch by cell index.
 */
function quantizePosition(pos: { x: number; y: number }, cellW: number, cellH: number): IconPos {
  const col = Math.max(0, Math.round((pos.x - DESKTOP_PAD_X) / cellW))
  const row = Math.max(0, Math.round((pos.y - DESKTOP_PAD_Y) / cellH))
  return { x: DESKTOP_PAD_X + col * cellW, y: DESKTOP_PAD_Y + row * cellH }
}

function quantizePositions(
  raw: Record<string, { x: number; y: number }> | undefined,
  cellW: number,
  cellH: number,
): Record<string, IconPos> {
  const out: Record<string, IconPos> = {}
  if (!raw) return out
  for (const [id, pos] of Object.entries(raw)) {
    if (!id || !pos || typeof pos.x !== 'number' || typeof pos.y !== 'number') continue
    out[id] = quantizePosition(pos, cellW, cellH)
  }
  return out
}

/** stable serialized form of a positions map, for change detection */
function serializePositions(positions: Readonly<Record<string, IconPos>>): string {
  return JSON.stringify(
    Object.keys(positions)
      .sort()
      .map((id) => [id, positions[id]!.x, positions[id]!.y]),
  )
}

/** the read-only sort flows pin nothing: every icon flows in sorted order */
const EMPTY_POSITIONS: Readonly<Record<string, IconPos>> = {}

/**
 * The Home desktop: an icon-desktop view that takes over the content area.
 * System app tiles, the folder roots and the newest recents lay out as a
 * column-major grid (top-left, down then right, like an OS desktop); positions
 * the person drags are saved to `<userData>/desktop-layout.json` (debounced,
 * through `desktopLayoutSave`) and restored on the next launch, while icons
 * without a saved spot flow into free cells without overlapping saved ones.
 *
 * The desktop's look-and-feel — wallpaper (and how it fills the surface),
 * timed carousel, icon size, sort mode, the label color — lives in
 * `<userData>/desktop-prefs.json` through `desktopPrefsSave`, applied the
 * moment its menu item is chosen. The right-click menu on empty surface
 * follows the OS conventions (新建 / 排序方式 / 查看 / 刷新 / 更换壁纸 /
 * 属性) with one level of flyout submenus; file and folder icons get a
 * properties dialog off their own menu. The wallpaper center and the two
 * properties dialogs render from their own files
 * (DesktopWallpaperCenter.tsx, DesktopProperties.tsx); this view owns the
 * state they share — the prefs, the wallpaper bytes, the open flags.
 *
 * The wallpaper backdrop itself renders one level up: Home hands the view the
 * `backdropHost` node it mounts at its root, and the backdrop portals there so
 * the whole Home window (sidebar included) sits on it, with the sidebar
 * wearing its glass treatment (desktop.css). The fit styles ride every
 * crossfade layer so a slideshow never disagrees with itself.
 *
 * Everything the icons do goes through `window.aiOffice` or one of the
 * parent's callbacks (ask view, template gallery, settings modal, folder
 * view) — the desktop owns only its own layout. The look lives in
 * desktop.css; the context menu reuses Home's `.folder-menu` conventions.
 */
export function DesktopView({
  lang,
  roots,
  onOpenAsk,
  onOpenTemplates,
  onOpenSettings,
  onOpenFolder,
  backdropHost,
}: DesktopViewProps): ReactElement {
  const s = desktopStrings(lang)

  // ── data ──
  const [recents, setRecents] = useState<readonly RecentEntry[]>([])

  // ── look-and-feel prefs (wallpaper, carousel, icon size, sort) ──
  const [prefs, setPrefs] = useState<DesktopPrefs>(DEFAULT_PREFS)

  // ── layout state ──
  /** saved positions only: what the file holds, what a drag or reset edits */
  const [positions, setPositions] = useState<Readonly<Record<string, IconPos>>>({})
  const [loaded, setLoaded] = useState(false)
  const [selected, setSelected] = useState<string | null>(null)
  const [drag, setDrag] = useState<{ id: string; x: number; y: number } | null>(null)
  const [menu, setMenu] = useState<{ id: string; x: number; y: number } | null>(null)

  // ── the empty-surface context menu and its one open submenu ──
  const [surfaceMenu, setSurfaceMenu] = useState<{ x: number; y: number } | null>(null)
  const [surfaceSub, setSurfaceSub] = useState<'new' | 'sort' | 'view' | null>(null)

  // ── the properties dialogs ──
  /** the icon whose properties dialog is open (file or folder icons only) */
  const [propsTarget, setPropsTarget] = useState<DesktopPropsTarget | null>(null)
  const [desktopPropsOpen, setDesktopPropsOpen] = useState(false)

  // ── the wallpaper center and the carousel ──
  const [wpCenterOpen, setWpCenterOpen] = useState(false)
  const [carouselIndex, setCarouselIndex] = useState(0)
  /** wallpaper data URLs by id, shared by the backdrop and the center's thumbnails */
  const [wallpaperImages, setWallpaperImages] = useState<Readonly<Record<string, string>>>({})
  /** the wallpaper ids currently mounted as crossfade layers (≤ 2, newest last) */
  const [wallpaperLayers, setWallpaperLayers] = useState<readonly string[]>([])
  const [reducedMotion] = useState(prefersReducedMotion)

  const surfaceRef = useRef<HTMLDivElement | null>(null)
  const menuRef = useRef<HTMLDivElement | null>(null)
  const surfaceMenuRef = useRef<HTMLDivElement | null>(null)
  const starfieldRef = useRef<HTMLCanvasElement | null>(null)
  /** the backdrop layer's own root (portaled or in place): what the starfield
   * canvas sizes itself against — the whole Home window in desktop mode */
  const backdropRef = useRef<HTMLDivElement | null>(null)
  const [surfaceSize, setSurfaceSize] = useState({ w: 0, h: 0 })

  // ── the icon set ──

  const icons = useMemo<DesktopIcon[]>(() => {
    const apps: DesktopIcon[] = [
      {
        id: 'app:docs',
        kind: 'app',
        tile: 'docs',
        label: s.appDocs,
        ext: '',
        path: null,
        run: () => void (window.aiOffice as Partial<HomeApi>).newDoc?.(),
      },
      {
        id: 'app:sheets',
        kind: 'app',
        tile: 'sheets',
        label: s.appSheets,
        ext: '',
        path: null,
        run: () => void (window.aiOffice as Partial<HomeApi>).newSheet?.(),
      },
      {
        id: 'app:slides',
        kind: 'app',
        tile: 'slides',
        label: s.appSlides,
        ext: '',
        path: null,
        run: () => void (window.aiOffice as Partial<HomeApi>).newSlide?.(),
      },
      {
        id: 'app:pdf',
        kind: 'app',
        tile: 'pdf',
        label: s.appPdf,
        ext: '',
        path: null,
        run: () => void (window.aiOffice as Partial<HomeApi>).newPdf?.(),
      },
      {
        id: 'app:markdown',
        kind: 'app',
        tile: 'markdown',
        label: s.appMarkdown,
        ext: '',
        path: null,
        run: () => void (window.aiOffice as Partial<HomeApi>).newMarkdown?.(),
      },
      {
        id: 'app:html',
        kind: 'app',
        tile: 'html',
        label: s.appHtml,
        ext: '',
        path: null,
        run: () => void (window.aiOffice as Partial<HomeApi>).newHtml?.(),
      },
      {
        id: 'app:ask',
        kind: 'app',
        tile: 'ask',
        label: s.appAsk,
        ext: '',
        path: null,
        run: onOpenAsk,
      },
      {
        id: 'app:templates',
        kind: 'app',
        tile: 'templates',
        label: s.appTemplates,
        ext: '',
        path: null,
        run: onOpenTemplates,
      },
      {
        id: 'app:plugins',
        kind: 'app',
        tile: 'plugins',
        label: s.appPlugins,
        ext: '',
        path: null,
        run: () => onOpenSettings({ section: 'plugins' }),
      },
      {
        id: 'app:settings',
        kind: 'app',
        tile: 'settings',
        label: s.appSettings,
        ext: '',
        path: null,
        run: () => onOpenSettings({ section: 'general' }),
      },
    ]
    const folders: DesktopIcon[] = roots.map((root) => ({
      id: `folder:${root.path}`,
      kind: 'folder',
      tile: null,
      label: root.name,
      ext: '',
      path: root.path,
      run: () => onOpenFolder(root),
    }))
    const files: DesktopIcon[] = recents.map((entry) => ({
      id: `file:${entry.path}`,
      kind: 'file',
      tile: null,
      label: entry.name,
      ext: entry.ext,
      path: entry.path,
      mtimeMs: entry.mtimeMs,
      run: () => void (window.aiOffice as Partial<HomeApi>).openPath?.(entry.path),
    }))
    return [...apps, ...folders, ...files]
  }, [s, roots, recents, onOpenAsk, onOpenTemplates, onOpenSettings, onOpenFolder])

  // mirrors of render-time values, read by the window-level event handlers
  const iconsRef = useRef(icons)
  iconsRef.current = icons
  const effectiveRef = useRef(new Map<string, IconPos>())
  const rowsRef = useRef(UNMEASURED_ROWS)
  const colsRef = useRef(1)

  // ── load: the saved layout, the prefs and the recents ──

  /** re-pull the newest recents (launch, a recents-changed push, the 刷新 item) */
  const reloadRecents = useCallback(() => {
    const request = (window.aiOffice as Partial<HomeApi>).recents?.({ limit: RECENT_LIMIT })
    if (!request) return
    request.then(
      (page) => setRecents(page.entries.filter((entry) => !entry.missing)),
      () => undefined,
    )
  }, [])

  useEffect(() => {
    let alive = true
    const api = window.aiOffice as Partial<HomeApi>
    const layoutRequest = api.desktopLayout?.()
    if (!layoutRequest) {
      savedRef.current = serializePositions({})
      setLoaded(true)
    } else {
      layoutRequest.then(
        (layout) => {
          if (!alive) return
          const loaded = quantizePositions(layout?.positions, cellRef.current.w, cellRef.current.h)
          // the loaded map is what is on disk: reloading it must not rewrite it
          savedRef.current = serializePositions(loaded)
          setPositions(loaded)
          setLoaded(true)
        },
        () => {
          if (alive) {
            savedRef.current = serializePositions({})
            setLoaded(true)
          }
        },
      )
    }
    const prefsRequest = api.desktopPrefs?.()
    if (prefsRequest) {
      prefsRequest.then(
        (loaded) => {
          if (alive) setPrefs(loaded)
        },
        () => undefined,
      )
    }
    reloadRecents()
    const off = api.onRecentsChanged?.(reloadRecents)
    return () => {
      alive = false
      off?.()
    }
  }, [reloadRecents])

  // ── write-through: every layout change arms one save, 300ms out ──

  const positionsRef = useRef(positions)
  positionsRef.current = positions
  /** the serialized map that is on disk or in flight; a flush of it is a no-op */
  const savedRef = useRef<string | null>(null)
  const savingRef = useRef(false)
  const retryTimerRef = useRef(0)

  useEffect(() => {
    if (!loaded) return
    const timer = window.setTimeout(() => void flush(), SAVE_DEBOUNCE_MS)
    return () => window.clearTimeout(timer)

    async function flush(): Promise<void> {
      if (savingRef.current) {
        retryTimerRef.current = window.setTimeout(() => void flush(), SAVE_DEBOUNCE_MS)
        return
      }
      const snapshot = serializePositions(positionsRef.current)
      if (snapshot === savedRef.current) return
      savingRef.current = true
      try {
        await (window.aiOffice as Partial<HomeApi>).desktopLayoutSave?.({
          positions: positionsRef.current,
        })
        savedRef.current = snapshot
        // a change landed mid-flight: write the newest map right away
        if (serializePositions(positionsRef.current) !== savedRef.current) {
          retryTimerRef.current = window.setTimeout(() => void flush(), 0)
        }
      } catch {
        // a refused save is retried by the next change's debounce
      } finally {
        savingRef.current = false
      }
    }
  }, [positions, loaded])

  useEffect(() => () => window.clearTimeout(retryTimerRef.current), [])

  // ── prefs write-through: one whole-object save per choice ──

  const prefsRef = useRef(prefs)
  prefsRef.current = prefs

  /** apply a prefs patch now and persist the whole object; a refused save is
   * retried by nothing (the next choice), the UI already shows the choice */
  const applyPrefs = useCallback((patch: Partial<DesktopPrefs>) => {
    const next = { ...prefsRef.current, ...patch }
    prefsRef.current = next
    setPrefs(next)
    const save = (window.aiOffice as Partial<HomeApi>).desktopPrefsSave
    if (save) void save(next).then(undefined, () => undefined)
  }, [])

  // ── the wallpaper: which id is showing, and its bytes ──

  /** the slideshow owns the backdrop while it is on, starting from the chosen
   * static when it is one; a motion-sensitive setup never rotates */
  const carouselActive = prefs.carouselMinutes > 0 && !reducedMotion
  const carouselCurrent = STATIC_WALLPAPERS[carouselIndex] ?? STATIC_WALLPAPERS[0]
  const effectiveWallpaper = carouselActive ? carouselCurrent : prefs.wallpaper
  /** '' when the effective wallpaper is the CSS aurora (no file behind it) */
  const wallpaperImageId = effectiveWallpaper === 'default' ? '' : effectiveWallpaper
  /** the fit row only exists while an image is actually showing: the live CSS
   * aurora has no image to fill, while the carousel always shows statics */
  const fitApplies = effectiveWallpaper !== 'default'

  /** a custom wallpaper's sampled darkness (average luminance < 128), one
   * canvas read per id — drives the automatic label ink over unknown art */
  const [customDark, setCustomDark] = useState<Record<string, boolean>>({})
  useEffect(() => {
    if (!effectiveWallpaper.startsWith('custom:')) return
    if (customDark[effectiveWallpaper] !== undefined) return
    const url = wallpaperImages[effectiveWallpaper]
    if (!url) return
    let alive = true
    const img = new Image()
    img.onload = () => {
      const c = document.createElement('canvas')
      c.width = 32
      c.height = 32
      const cx = c.getContext('2d')
      if (!cx) return
      cx.drawImage(img, 0, 0, 32, 32)
      const d = cx.getImageData(0, 0, 32, 32).data
      let sum = 0
      for (let i = 0; i < d.length; i += 4) {
        sum += 0.2126 * d[i]! + 0.7152 * d[i + 1]! + 0.0722 * d[i + 2]!
      }
      if (alive) setCustomDark((m) => ({ ...m, [effectiveWallpaper]: sum / (d.length / 4) < 128 }))
    }
    img.src = url
    return () => {
      alive = false
    }
  }, [effectiveWallpaper, wallpaperImages, customDark])

  /** the aurora is theme-adaptive; image wallpapers carry their own darkness */
  const darkChrome =
    typeof document !== 'undefined' &&
    (document.documentElement.dataset.theme === 'dark' ||
      ((!document.documentElement.dataset.theme ||
        document.documentElement.dataset.theme === 'system') &&
        typeof window.matchMedia === 'function' &&
        window.matchMedia('(prefers-color-scheme: dark)').matches))
  const darkWallpaper =
    effectiveWallpaper === 'default'
      ? darkChrome
      : DARK_WALLPAPERS.has(effectiveWallpaper)
        ? true
        : effectiveWallpaper.startsWith('custom:')
          ? (customDark[effectiveWallpaper] ?? false)
          : false

  // the index follows the chosen static, so turning the slideshow on starts
  // it where the person already is
  useEffect(() => {
    const at = STATIC_WALLPAPERS.indexOf(prefs.wallpaper as (typeof STATIC_WALLPAPERS)[number])
    setCarouselIndex(at >= 0 ? at : 0)
  }, [prefs.wallpaper])

  useEffect(() => {
    if (!carouselActive) return
    const timer = window.setInterval(
      () => setCarouselIndex((i) => (i + 1) % STATIC_WALLPAPERS.length),
      prefs.carouselMinutes * 60_000,
    )
    return () => window.clearInterval(timer)
  }, [carouselActive, prefs.carouselMinutes])

  // the crossfade layer stack: a change keeps the previous image mounted and
  // fading out under the new one (opacity only, never a transform)
  useEffect(() => {
    setWallpaperLayers((prev) => {
      if (prev[prev.length - 1] === wallpaperImageId) return prev
      return [...prev, wallpaperImageId].slice(-2)
    })
  }, [wallpaperImageId])

  /** every id the view needs bytes for: the backdrop's, the fading-out one,
   * and the whole catalog while the wallpaper center is open */
  const neededWallpaperIds = useMemo(() => {
    const ids = new Set<string>()
    if (wallpaperImageId !== '') ids.add(wallpaperImageId)
    for (const id of wallpaperLayers) if (id !== '') ids.add(id)
    if (wpCenterOpen) {
      for (const id of ANIMATED_WALLPAPERS) if (id !== 'default') ids.add(id)
      for (const id of STATIC_WALLPAPERS) ids.add(id)
    }
    return [...ids]
  }, [wallpaperImageId, wallpaperLayers, wpCenterOpen])

  const wallpaperImagesRef = useRef(wallpaperImages)
  wallpaperImagesRef.current = wallpaperImages
  const wallpaperPendingRef = useRef(new Set<string>())

  // fetch each needed id once; failures are dropped so a retry can happen
  useEffect(() => {
    const api = window.aiOffice as Partial<HomeApi>
    if (!api.wallpaperImage) return
    for (const id of neededWallpaperIds) {
      if (wallpaperImagesRef.current[id] !== undefined) continue
      if (wallpaperPendingRef.current.has(id)) continue
      wallpaperPendingRef.current.add(id)
      void api.wallpaperImage(id).then(
        (url) => {
          wallpaperPendingRef.current.delete(id)
          if (url !== '') {
            setWallpaperImages((prev) => (id in prev ? prev : { ...prev, [id]: url }))
          }
        },
        () => {
          wallpaperPendingRef.current.delete(id)
        },
      )
    }
  }, [neededWallpaperIds])

  // the starfield animation layer: twinkling stars over the static PNG.
  // Nothing draws without a 2d context (test DOMs), and a motion-sensitive
  // setup never mounts it at all — the PNG's own stars stand still.
  const showStarfield = effectiveWallpaper === 'builtin:starfield-night' && !reducedMotion
  useEffect(() => {
    if (!showStarfield) return
    const canvas = starfieldRef.current
    const backdrop = backdropRef.current
    if (!canvas || !backdrop) return
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    interface DriftStar {
      x: number
      y: number
      r: number
      phase: number
      speed: number
    }
    const stars: DriftStar[] = []
    const resize = () => {
      const rect = backdrop.getBoundingClientRect()
      canvas.width = Math.max(1, Math.round(rect.width))
      canvas.height = Math.max(1, Math.round(rect.height))
      stars.length = 0
      const count = Math.max(24, Math.min(160, Math.round((canvas.width * canvas.height) / 14000)))
      for (let i = 0; i < count; i += 1) {
        stars.push({
          x: Math.random() * canvas.width,
          y: Math.random() * canvas.height,
          r: 0.6 + Math.random() * 1.4,
          phase: Math.random() * Math.PI * 2,
          speed: 0.4 + Math.random() * 0.8,
        })
      }
    }
    resize()
    window.addEventListener('resize', resize)
    let frame = 0
    let tick = 0
    const draw = (t: number) => {
      // the twinkle is a slow sine (t * 0.001) — drawing it at 12fps is
      // visually identical to 60fps and costs a fifth of the GPU/CPU
      tick += 1
      if (tick % 5 === 0) {
        ctx.clearRect(0, 0, canvas.width, canvas.height)
        for (const star of stars) {
          const twinkle = 0.5 + 0.5 * Math.sin(t * 0.001 * star.speed + star.phase)
          const alpha = 0.22 + 0.38 * twinkle
          ctx.fillStyle = `rgba(226, 236, 255, ${alpha.toFixed(3)})`
          ctx.beginPath()
          ctx.arc(star.x, star.y, star.r, 0, Math.PI * 2)
          ctx.fill()
        }
      }
      frame = requestAnimationFrame(draw)
    }
    frame = requestAnimationFrame(draw)
    return () => {
      cancelAnimationFrame(frame)
      window.removeEventListener('resize', resize)
    }
  }, [showStarfield])

  // ── surface measurement: the grid's row count depends on the height ──

  useLayoutEffect(() => {
    const measure = () => {
      const el = surfaceRef.current
      if (!el) return
      const rect = el.getBoundingClientRect()
      setSurfaceSize((prev) =>
        prev.w === rect.width && prev.h === rect.height ? prev : { w: rect.width, h: rect.height },
      )
    }
    measure()
    window.addEventListener('resize', measure)
    return () => window.removeEventListener('resize', measure)
  }, [])

  // ── the cell pitch, scaled by the icon size (查看 ▸ 图标大小) ──
  // The same multiplier desktop.css applies to the tiles, so a cell always
  // fits its contents; saved positions re-quantize onto the same row/column.

  const iconScale = ICON_SCALES[prefs.iconSize]
  const cellW = Math.round(DESKTOP_CELL_W * iconScale)
  const cellH = Math.round(DESKTOP_CELL_H * iconScale)
  const cellRef = useRef({ w: cellW, h: cellH })
  cellRef.current = { w: cellW, h: cellH }

  const rows =
    surfaceSize.h > 0
      ? Math.max(1, Math.floor((surfaceSize.h - DESKTOP_PAD_Y * 2) / cellH))
      : UNMEASURED_ROWS
  const cols =
    surfaceSize.w > 0 ? Math.max(1, Math.floor((surfaceSize.w - DESKTOP_PAD_X * 2) / cellW)) : 1
  rowsRef.current = rows
  colsRef.current = cols

  // ── the resolved grid: saved positions pin, the rest flow into free cells ──

  /** the read-only sort flows reorder the icon set and ignore the saved map;
   * choosing 自定义 restores it untouched */
  const orderedIcons = useMemo(
    () => sortDesktopIcons(icons, prefs.sortMode),
    [icons, prefs.sortMode],
  )
  const freeLayout = prefs.sortMode === 'custom'
  const layoutPositions = freeLayout ? positions : EMPTY_POSITIONS

  const resolved = useMemo(() => {
    const cellOf = new Map<string, { col: number; row: number }>()
    const taken = new Set<string>()
    // pinned: quantized saved cells that are still inside the visible grid.
    // Two icons pinned to one cell (a hand-edited file, a shrunken window)
    // keep the first; the later one falls back to the free-cell flow.
    for (const icon of orderedIcons) {
      const pos = layoutPositions[icon.id]
      if (!pos) continue
      const cell = quantizePosition(pos, cellW, cellH)
      const col = (cell.x - DESKTOP_PAD_X) / cellW
      const row = (cell.y - DESKTOP_PAD_Y) / cellH
      if (col >= cols || row >= rows) continue
      const key = `${col}:${row}`
      if (taken.has(key)) continue
      taken.add(key)
      cellOf.set(icon.id, { col, row })
    }
    // free icons flow in column-major order: top-left, down, then the next column
    for (const icon of orderedIcons) {
      if (cellOf.has(icon.id)) continue
      for (let col = 0; ; col += 1) {
        let placed = false
        for (let row = 0; row < rows; row += 1) {
          const key = `${col}:${row}`
          if (taken.has(key)) continue
          taken.add(key)
          cellOf.set(icon.id, { col, row })
          placed = true
          break
        }
        if (placed) break
      }
    }
    const out = orderedIcons.map((icon) => {
      const cell = cellOf.get(icon.id) ?? { col: 0, row: 0 }
      return {
        icon,
        pos: {
          x: DESKTOP_PAD_X + cell.col * cellW,
          y: DESKTOP_PAD_Y + cell.row * cellH,
        },
      }
    })
    effectiveRef.current = new Map(out.map(({ icon, pos }) => [icon.id, pos]))
    return out
  }, [orderedIcons, layoutPositions, rows, cols, cellW, cellH])

  // ── dragging (pointer events): the icon follows the pointer, drops snap ──
  // A press on an icon arms one effect-scoped set of window listeners for the
  // whole gesture, so every handler closes over the press it serves and no
  // callback ever goes stale mid-drag.

  interface PressState {
    id: string
    pointerId: number
    grabX: number
    grabY: number
    startX: number
    startY: number
  }
  const [press, setPress] = useState<PressState | null>(null)
  /** last drag end; a double-click fired this soon after is not an open */
  const dropGuardRef = useRef(0)

  /** commit a drop: the dragged icon claims the snapped cell, an occupant swaps out */
  const commitDrop = useCallback((id: string, x: number, y: number) => {
    const { w: pitchW, h: pitchH } = cellRef.current
    const col = clampInt(Math.round((x - DESKTOP_PAD_X) / pitchW), 0, colsRef.current - 1)
    const row = clampInt(Math.round((y - DESKTOP_PAD_Y) / pitchH), 0, rowsRef.current - 1)
    const target = {
      x: DESKTOP_PAD_X + col * pitchW,
      y: DESKTOP_PAD_Y + row * pitchH,
    }
    const from = effectiveRef.current.get(id)
    if (!from || (from.x === target.x && from.y === target.y)) return
    // whoever sits at the target cell takes the dragged icon's old cell
    let occupant: string | null = null
    for (const [otherId, pos] of effectiveRef.current) {
      if (otherId !== id && pos.x === target.x && pos.y === target.y) {
        occupant = otherId
        break
      }
    }
    setPositions((prev) => {
      const next: Record<string, IconPos> = { ...prev }
      next[id] = target
      if (occupant !== null) next[occupant] = from
      return next
    })
  }, [])

  useEffect(() => {
    if (press === null) return
    const d = press
    let moved = false
    let last: { x: number; y: number } | null = null
    const onMove = (event: PointerEvent) => {
      if (pointerIdOf(event) !== d.pointerId) return
      const surface = surfaceRef.current
      if (!surface) return
      if (!moved) {
        if (Math.hypot(event.clientX - d.startX, event.clientY - d.startY) < DRAG_THRESHOLD) return
        moved = true
      }
      const rect = surface.getBoundingClientRect()
      last = {
        x: Math.max(0, event.clientX - rect.left - d.grabX),
        y: Math.max(0, event.clientY - rect.top - d.grabY),
      }
      setDrag({ id: d.id, x: last.x, y: last.y })
    }
    const finish = (commit: boolean) => {
      if (commit && moved && last) {
        dropGuardRef.current = Date.now() + DROP_DBLCLICK_GUARD_MS
        commitDrop(d.id, last.x, last.y)
      }
      setPress(null)
      setDrag(null)
    }
    const onUp = (event: PointerEvent) => {
      if (pointerIdOf(event) !== d.pointerId) return
      finish(true)
    }
    // a cancel (or the window losing focus mid-drag) puts the icon back untouched
    const onCancel = () => finish(false)
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
    window.addEventListener('pointercancel', onCancel)
    window.addEventListener('blur', onCancel)
    return () => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
      window.removeEventListener('pointercancel', onCancel)
      window.removeEventListener('blur', onCancel)
    }
  }, [press, commitDrop])

  const onIconPointerDown = useCallback(
    (event: ReactPointerEvent, icon: DesktopIcon, pos: IconPos) => {
      if (event.button !== 0 || press !== null) return
      const surface = surfaceRef.current
      if (!surface) return
      setSelected(icon.id)
      setMenu(null)
      setSurfaceMenu(null)
      // the read-only sort flows are look-don't-touch: a press selects, but
      // never arms the drag machinery (自定义 is the only draggable mode)
      if (!freeLayout) return
      const rect = surface.getBoundingClientRect()
      setPress({
        id: icon.id,
        pointerId: pointerIdOf(event),
        grabX: event.clientX - rect.left - pos.x,
        grabY: event.clientY - rect.top - pos.y,
        startX: event.clientX,
        startY: event.clientY,
      })
    },
    [press, freeLayout],
  )

  const openIcon = useCallback((id: string) => {
    iconsRef.current.find((icon) => icon.id === id)?.run()
  }, [])

  // ── context menus (Home's .folder-menu conventions) ──

  useDismissablePopover(menu !== null, () => setMenu(null), {
    inside: () => [menuRef.current],
  })

  useDismissablePopover(
    surfaceMenu !== null,
    () => {
      setSurfaceMenu(null)
      setSurfaceSub(null)
    },
    { inside: () => [surfaceMenuRef.current] },
  )

  useLayoutEffect(() => {
    const el = menuRef.current
    if (!menu || !el) return
    const { left, top } = placeFolderMenu(
      { x: menu.x, y: menu.y, above: menu.y, align: 'start' },
      { width: el.offsetWidth, height: el.offsetHeight },
      {
        width: document.documentElement.clientWidth,
        height: document.documentElement.clientHeight,
      },
    )
    el.style.left = `${left}px`
    el.style.top = `${top}px`
  }, [menu])

  useLayoutEffect(() => {
    const el = surfaceMenuRef.current
    if (!surfaceMenu || !el) return
    const { left, top } = placeFolderMenu(
      { x: surfaceMenu.x, y: surfaceMenu.y, above: surfaceMenu.y, align: 'start' },
      { width: el.offsetWidth, height: el.offsetHeight },
      {
        width: document.documentElement.clientWidth,
        height: document.documentElement.clientHeight,
      },
    )
    el.style.left = `${left}px`
    el.style.top = `${top}px`
    // flyouts open to the right of their row; near the viewport's right edge
    // they mirror to the left so nothing is cut off
    el.classList.toggle(
      'flip-sub',
      left + el.offsetWidth + 190 > document.documentElement.clientWidth,
    )
  }, [surfaceMenu])

  const menuIcon = menu ? icons.find((icon) => icon.id === menu.id) : undefined
  const menuPath = menuIcon?.kind === 'file' ? menuIcon.path : null
  /** file and folder icons get a 属性 dialog; app icons do not */
  const menuPropsIcon = menuIcon && menuIcon.kind !== 'app' ? menuIcon : null

  const resetPosition = useCallback((id: string) => {
    setMenu(null)
    setPositions((prev) => {
      if (!(id in prev)) return prev
      const next: Record<string, IconPos> = { ...prev }
      delete next[id]
      return next
    })
  }, [])

  // ── surface-menu actions ──

  /** run one of the 新建 create calls and close the menu */
  const createDocument = useCallback((action: (typeof NEW_DOCUMENT_ACTIONS)[number]['action']) => {
    setSurfaceMenu(null)
    setSurfaceSub(null)
    void (window.aiOffice as Partial<HomeApi>)[action]?.()
  }, [])

  const refreshRecents = useCallback(() => {
    setSurfaceMenu(null)
    setSurfaceSub(null)
    reloadRecents()
  }, [reloadRecents])

  const openWallpaperCenter = useCallback(() => {
    setSurfaceMenu(null)
    setSurfaceSub(null)
    setWpCenterOpen(true)
  }, [])

  const openDesktopProps = useCallback(() => {
    setSurfaceMenu(null)
    setSurfaceSub(null)
    setDesktopPropsOpen(true)
  }, [])

  /** open the icon properties dialog off the icon's own menu */
  const openIconProps = useCallback((icon: DesktopIcon) => {
    setMenu(null)
    if (icon.path === null) return
    setPropsTarget({
      kind: icon.kind === 'folder' ? 'folder' : 'file',
      label: icon.label,
      ext: icon.ext,
      path: icon.path,
    })
  }, [])

  /** apply a picked wallpaper and close the center — the backdrop re-renders
   * behind the dialog on its way out */
  const applyWallpaper = useCallback(
    (id: string) => {
      setWpCenterOpen(false)
      applyPrefs({ wallpaper: id })
    },
    [applyPrefs],
  )

  // ── keyboard: Enter opens the selection, Escape closes menu/selection,
  // arrows move it across the grid, typing jumps to a matching label ──
  // (a dialog, while open, owns Escape itself)

  const dialogOpen = propsTarget !== null || desktopPropsOpen || wpCenterOpen

  /** the type-ahead buffer: keystrokes within 800ms extend the search */
  const typeAheadRef = useRef({ text: '', at: 0 })

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        if (dialogOpen) return
        if (surfaceMenu !== null) setSurfaceMenu(null)
        else if (menu !== null) setMenu(null)
        else if (selected !== null) setSelected(null)
        return
      }
      if (event.metaKey || event.ctrlKey || event.altKey) return
      if (dialogOpen) return
      // only when the desktop owns the keyboard: an icon, the surface, or nothing focused
      const target = event.target
      const owns =
        !(target instanceof Element) ||
        target === document.body ||
        (surfaceRef.current?.contains(target) ?? false)
      if (!owns) return

      // grid navigation follows the visible column-major flow: down/up step
      // inside a column, right/left jump by a whole column
      const move = (delta: number) => {
        const list = orderedIcons
        if (list.length === 0) return
        const at = selected === null ? -1 : list.findIndex((i) => i.id === selected)
        const next =
          at === -1 ? (delta > 0 ? 0 : list.length - 1) : (at + delta + list.length) % list.length
        setSelected(list[next]!.id)
      }
      if (event.key === 'ArrowDown') {
        event.preventDefault()
        move(1)
        return
      }
      if (event.key === 'ArrowUp') {
        event.preventDefault()
        move(-1)
        return
      }
      if (event.key === 'ArrowRight') {
        event.preventDefault()
        move(rowsRef.current)
        return
      }
      if (event.key === 'ArrowLeft') {
        event.preventDefault()
        move(-rowsRef.current)
        return
      }

      if (event.key === 'Enter') {
        if (event.shiftKey || selected === null) return
        event.preventDefault()
        openIcon(selected)
        return
      }

      // type-ahead: printable runs within 800ms jump to the first label match
      if (event.key.length === 1) {
        const now = Date.now()
        const buf = now - typeAheadRef.current.at > 800 ? '' : typeAheadRef.current.text
        const next = (buf + event.key).toLowerCase()
        typeAheadRef.current = { text: next, at: now }
        const hit = orderedIcons.find((i) => i.label.toLowerCase().startsWith(next))
        if (hit) {
          event.preventDefault()
          setSelected(hit.id)
        }
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [menu, surfaceMenu, selected, openIcon, dialogOpen, orderedIcons])

  // the dialogs' own Escape: the topmost one closes
  useEffect(() => {
    if (!dialogOpen) return
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return
      if (wpCenterOpen) setWpCenterOpen(false)
      else if (propsTarget !== null) setPropsTarget(null)
      else if (desktopPropsOpen) setDesktopPropsOpen(false)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [dialogOpen, wpCenterOpen, propsTarget, desktopPropsOpen])

  // ── render ──

  /** the label color's two custom properties; a custom color always wins, a
   * dark wallpaper under '' flips the automatic ink to light (the theme pair
   * stays in charge over light art) */
  const labelStyle: CSSProperties | undefined = useMemo(() => {
    const ink = prefs.labelColor !== '' ? prefs.labelColor : darkWallpaper ? '#ffffff' : ''
    if (ink === '') return undefined
    return {
      '--desktop-label-ink': ink,
      '--desktop-label-scrim': labelScrimOf(ink),
    } as CSSProperties
  }, [prefs.labelColor, darkWallpaper])

  const renderTile = (icon: DesktopIcon): ReactNode => {
    if (icon.kind === 'file') return <FileBadge ext={icon.ext} size={64} />
    if (icon.kind === 'folder') {
      return (
        <span className="desktop-tile desktop-tile-folder">
          <FolderGlyph />
        </span>
      )
    }
    return (
      <span className={`desktop-tile desktop-tile-${icon.tile}`}>
        {icon.tile === 'docs' && <DocsGlyph />}
        {icon.tile === 'sheets' && <SheetsGlyph />}
        {icon.tile === 'slides' && <SlidesGlyph />}
        {icon.tile === 'pdf' && <PdfGlyph />}
        {icon.tile === 'markdown' && <MarkdownGlyph />}
        {icon.tile === 'html' && <HtmlGlyph />}
        {icon.tile === 'ask' && <SparkGlyph />}
        {icon.tile === 'templates' && <TemplatesGlyph />}
        {icon.tile === 'plugins' && <PlugGlyph />}
        {icon.tile === 'settings' && <GearGlyph />}
      </span>
    )
  }

  /** the wallpaper backdrop: lifted to the Home root through the portal host
   * Home provides (so the sidebar sits on it too), or rendered in place when
   * the view stands alone. Null while Home's host node is committing. */
  const backdrop = (
    <div className="desktop-backdrop" ref={backdropRef} aria-hidden="true">
      <div className="desktop-wallpaper">
        {wallpaperLayers
          .filter((id) => id !== '' && wallpaperImages[id] !== undefined)
          .map((id) => {
            const src = wallpaperImages[id]
            if (src === undefined) return null
            return (
              <div
                key={id}
                className={id === wallpaperImageId ? 'active' : undefined}
                style={{
                  backgroundImage: `url("${src}")`,
                  ...WALLPAPER_FIT_STYLE[prefs.wallpaperFit],
                }}
              />
            )
          })}
      </div>
      {effectiveWallpaper === 'default' && (
        <>
          <div className="desktop-aurora desktop-aurora-a" />
          <div className="desktop-aurora desktop-aurora-b" />
          <div className="desktop-aurora desktop-aurora-c" />
          <div className="desktop-dots" />
        </>
      )}
      {showStarfield && <canvas ref={starfieldRef} className="desktop-starfield" />}
    </div>
  )

  return (
    <main className={`content desktop-content icons-${prefs.iconSize}`} style={labelStyle}>
      {backdropHost
        ? createPortal(backdrop, backdropHost)
        : backdropHost === undefined
          ? backdrop
          : null}
      <div
        className="desktop-surface"
        ref={surfaceRef}
        onPointerDown={(event) => {
          if (event.target !== event.currentTarget) return
          setSelected(null)
          setMenu(null)
          setSurfaceMenu(null)
          setSurfaceSub(null)
        }}
        onContextMenu={(event) => {
          // only empty surface: an icon's own menu wins on the icon itself
          if (event.target !== event.currentTarget) return
          event.preventDefault()
          setSelected(null)
          setMenu(null)
          setSurfaceMenu({ x: event.clientX + 2, y: event.clientY + 2 })
          setSurfaceSub(null)
        }}
      >
        <ul className="desktop-icons" role="listbox" aria-label={s.desktopAria}>
          {resolved.map(({ icon, pos }, index) => {
            const dragging = drag !== null && drag.id === icon.id
            const style = dragging
              ? { left: drag.x, top: drag.y, zIndex: 40 }
              : { left: pos.x, top: pos.y }
            return (
              <li
                key={icon.id}
                className={`desktop-icon${dragging ? ' dragging' : ''}${selected === icon.id ? ' selected' : ''}`}
                style={{ ...style, animationDelay: `${Math.min(index, 24) * 60}ms` }}
                role="presentation"
              >
                <div
                  className="desktop-icon-item"
                  role="option"
                  aria-selected={selected === icon.id}
                  tabIndex={0}
                  aria-label={icon.label}
                  title={icon.kind === 'folder' ? (icon.path ?? icon.label) : icon.label}
                  onPointerDown={(event) => onIconPointerDown(event, icon, pos)}
                  onDoubleClick={() => {
                    if (Date.now() < dropGuardRef.current) return
                    icon.run()
                  }}
                  onContextMenu={(event) => {
                    event.preventDefault()
                    setSelected(icon.id)
                    setSurfaceMenu(null)
                    setMenu({ id: icon.id, x: event.clientX + 2, y: event.clientY + 2 })
                  }}
                >
                  {renderTile(icon)}
                  <span className="desktop-icon-label">{icon.label}</span>
                </div>
              </li>
            )
          })}
        </ul>
        {menu && (
          <div
            className="folder-menu"
            role="menu"
            ref={(el) => {
              menuRef.current = el
            }}
          >
            <button
              role="menuitem"
              onClick={() => {
                setMenu(null)
                openIcon(menu.id)
              }}
            >
              {s.menuOpen}
            </button>
            {menuPath !== null && (
              <button
                role="menuitem"
                onClick={() => {
                  setMenu(null)
                  if (menuPath !== null) {
                    void (window.aiOffice as Partial<HomeApi>).revealPath?.(menuPath)
                  }
                }}
              >
                {s.menuReveal}
              </button>
            )}
            {menuPropsIcon !== null && (
              <button role="menuitem" onClick={() => openIconProps(menuPropsIcon)}>
                {s.menuProps}
              </button>
            )}
            <div className="row-menu-divider" />
            <button role="menuitem" onClick={() => resetPosition(menu.id)}>
              {s.menuReset}
            </button>
          </div>
        )}
        {surfaceMenu && (
          <div
            className="folder-menu desktop-menu"
            role="menu"
            aria-label={s.navDesktop}
            ref={(el) => {
              surfaceMenuRef.current = el
            }}
          >
            <div className="desktop-menu-row">
              <button
                role="menuitem"
                aria-haspopup="menu"
                aria-expanded={surfaceSub === 'new'}
                onClick={() => setSurfaceSub(surfaceSub === 'new' ? null : 'new')}
              >
                {s.menuNew}
              </button>
              {surfaceSub === 'new' && (
                <div className="desktop-submenu" role="menu">
                  {NEW_DOCUMENT_ACTIONS.map(({ action, labelKey }) => (
                    <button key={action} role="menuitem" onClick={() => createDocument(action)}>
                      {s[labelKey]}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <div className="desktop-menu-row">
              <button
                role="menuitem"
                aria-haspopup="menu"
                aria-expanded={surfaceSub === 'sort'}
                onClick={() => setSurfaceSub(surfaceSub === 'sort' ? null : 'sort')}
              >
                {s.menuSort}
              </button>
              {surfaceSub === 'sort' && (
                <div className="desktop-submenu" role="menu">
                  {SORT_MODE_ITEMS.map(({ mode, labelKey }) => (
                    <button
                      key={mode}
                      role="menuitem"
                      className={prefs.sortMode === mode ? 'checked' : undefined}
                      onClick={() => {
                        setSurfaceMenu(null)
                        setSurfaceSub(null)
                        applyPrefs({ sortMode: mode })
                      }}
                    >
                      {s[labelKey]}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <div className="desktop-menu-row">
              <button
                role="menuitem"
                aria-haspopup="menu"
                aria-expanded={surfaceSub === 'view'}
                onClick={() => setSurfaceSub(surfaceSub === 'view' ? null : 'view')}
              >
                {s.menuView}
              </button>
              {surfaceSub === 'view' && (
                <div className="desktop-submenu" role="menu">
                  <div className="desktop-submenu-label">{s.iconSizeLabel}</div>
                  {ICON_SIZE_ITEMS.map(({ size, labelKey }) => (
                    <button
                      key={size}
                      role="menuitem"
                      className={prefs.iconSize === size ? 'checked' : undefined}
                      onClick={() => {
                        setSurfaceMenu(null)
                        setSurfaceSub(null)
                        applyPrefs({ iconSize: size })
                      }}
                    >
                      {s[labelKey]}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <div className="row-menu-divider" />
            <button role="menuitem" onClick={refreshRecents}>
              {s.menuRefresh}
            </button>
            <button role="menuitem" onClick={openWallpaperCenter}>
              {s.menuWallpaper}
            </button>
            <div className="row-menu-divider" />
            <button role="menuitem" onClick={openDesktopProps}>
              {s.menuProps}
            </button>
          </div>
        )}
      </div>
      {propsTarget !== null && (
        <IconPropsDialog
          key={propsTarget.path}
          target={propsTarget}
          lang={lang}
          strings={s}
          onClose={() => setPropsTarget(null)}
        />
      )}
      {desktopPropsOpen && (
        <DesktopPropsDialog
          strings={s}
          fileCount={icons.filter((icon) => icon.kind === 'file').length}
          folderCount={icons.filter((icon) => icon.kind === 'folder').length}
          onClose={() => setDesktopPropsOpen(false)}
          onOpenSettings={() => onOpenSettings({ section: 'general' })}
        />
      )}
      {wpCenterOpen && (
        <DesktopWallpaperCenter
          strings={s}
          prefs={prefs}
          applyPrefs={applyPrefs}
          applyWallpaper={applyWallpaper}
          wallpaperImages={wallpaperImages}
          fitApplies={fitApplies}
          onClose={() => setWpCenterOpen(false)}
        />
      )}
    </main>
  )
}

// ── glyphs ──────────────────────────────────────────────

function FolderGlyph() {
  return (
    <svg width="34" height="34" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <path
        d="M1.5 4A1.5 1.5 0 0 1 3 2.5h3.1c.44 0 .85.19 1.13.52L8.4 4.4H13A1.5 1.5 0 0 1 14.5 5.9v5.6A1.5 1.5 0 0 1 13 13H3a1.5 1.5 0 0 1-1.5-1.5V4z"
        stroke="currentColor"
        strokeWidth="1.2"
        strokeLinejoin="round"
        fill="currentColor"
        fillOpacity="0.14"
      />
    </svg>
  )
}

function DocsGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <rect x="6.5" y="3.5" width="19" height="25" rx="3" stroke="currentColor" strokeWidth="2.4" />
      <path
        d="M11 11h10M11 16h10M11 21h6"
        stroke="currentColor"
        strokeWidth="2.4"
        strokeLinecap="round"
      />
    </svg>
  )
}

function SheetsGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <rect x="4.5" y="4.5" width="23" height="23" rx="3" stroke="currentColor" strokeWidth="2.2" />
      <path d="M4.5 11.5h23M11.5 11.5v16" stroke="currentColor" strokeWidth="2.2" />
      <rect x="13.5" y="13.5" width="6" height="5" rx="1" fill="currentColor" opacity="0.55" />
      <rect x="20.5" y="19.5" width="5" height="6" rx="1" fill="currentColor" opacity="0.55" />
    </svg>
  )
}

function SlidesGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <rect
        x="3.5"
        y="5.5"
        width="25"
        height="16"
        rx="2.5"
        stroke="currentColor"
        strokeWidth="2.4"
      />
      <path d="M13 10.5l6 3.5-6 3.5z" fill="currentColor" />
      <path
        d="M12 27.5h8M16 21.5v6"
        stroke="currentColor"
        strokeWidth="2.4"
        strokeLinecap="round"
      />
    </svg>
  )
}

function PdfGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <rect x="6.5" y="3.5" width="19" height="25" rx="3" stroke="currentColor" strokeWidth="2.4" />
      <text
        x="16"
        y="17.5"
        textAnchor="middle"
        dominantBaseline="central"
        fill="currentColor"
        fontSize="8.5"
        fontWeight="800"
        fontFamily="inherit"
        letterSpacing="0.5"
      >
        PDF
      </text>
    </svg>
  )
}

function MarkdownGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <rect x="4.5" y="6.5" width="23" height="19" rx="3" stroke="currentColor" strokeWidth="2.4" />
      <path
        d="M9 20v-7l3.5 4 3.5-4v7"
        stroke="currentColor"
        strokeWidth="2.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <path
        d="M21 13v7M18.5 17l2.5 3 2.5-3"
        stroke="currentColor"
        strokeWidth="2.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function HtmlGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <path
        d="M11.5 10.5L7 16l4.5 5.5M20.5 10.5L25 16l-4.5 5.5"
        stroke="currentColor"
        strokeWidth="2.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function SparkGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <path
        d="M16 3.5l3.1 9.4 9.4 3.1-9.4 3.1-3.1 9.4-3.1-9.4-9.4-3.1 9.4-3.1z"
        fill="currentColor"
      />
    </svg>
  )
}

function TemplatesGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" fill="none" aria-hidden="true">
      <rect
        x="9.5"
        y="4.5"
        width="18"
        height="23"
        rx="2.5"
        stroke="currentColor"
        strokeWidth="2.4"
        opacity="0.55"
      />
      <rect x="4.5" y="8.5" width="18" height="19" rx="2.5" fill="currentColor" opacity="0.25" />
      <rect
        x="4.5"
        y="8.5"
        width="18"
        height="19"
        rx="2.5"
        stroke="currentColor"
        strokeWidth="2.4"
      />
      <path
        d="M9 14h9M9 18.5h9M9 23h5"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
      />
    </svg>
  )
}

function PlugGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M9 2.5v5.5M15 2.5v5.5M6 8h12v4a6 6 0 0 1-6 6 6 6 0 0 1-6-6V8zM12 18v3.5"
        stroke="currentColor"
        strokeWidth="2.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function GearGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle cx="12" cy="12" r="3.2" stroke="currentColor" strokeWidth="2.2" />
      <path
        d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinejoin="round"
      />
    </svg>
  )
}
