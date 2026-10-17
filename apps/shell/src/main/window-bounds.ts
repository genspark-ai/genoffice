/**
 * The shell window's geometry across sessions (issue genoffice#2049) — the
 * position and size the app reopens with.
 *
 * The state lives in userData/app-settings.json, the file the issue names,
 * under one key. What is stored is `getNormalBounds()` — the window's
 * normal-state rectangle — never the live bounds: a maximized, minimized or
 * fullscreen window would otherwise come back as a screen-sized ordinary
 * window. The maximized flag is stored beside the rectangle and re-applied at
 * launch.
 *
 * Reading treats the file as untrusted (it is a plain JSON file a user can
 * edit): malformed geometry falls back to the default placement, and so does
 * a saved position no current display can show — a window whose monitor was
 * unplugged must not open where it cannot be grabbed. The title strip (the
 * tab bar) is what a user grabs to move the window, so that strip is what has
 * to land on some display's work area.
 */
import { readAppSettings, writeAppSetting } from './app-settings'

/** `app-settings.json` key holding the last session's window state. */
export const WINDOW_BOUNDS_KEY = 'windowBounds'

/**
 * Constructor values for the shell window (index.ts). They live here so the
 * creation defaults and the validation bounds cannot drift apart.
 */
export const SHELL_WINDOW_DEFAULT_WIDTH = 1360
export const SHELL_WINDOW_DEFAULT_HEIGHT = 900
export const SHELL_WINDOW_MIN_WIDTH = 720
export const SHELL_WINDOW_MIN_HEIGHT = 550

/**
 * A saved dimension beyond this is corrupt — the largest real displays are
 * 8K, and a giant window is worse than the default placement.
 */
const MAX_DIMENSION = 8192

/** How much of the title strip must show on a display for the window to count as reachable. */
const MIN_VISIBLE_WIDTH = 120
const TITLE_STRIP_HEIGHT = 40

/**
 * How long a drag or resize must settle before the geometry is written. Long
 * enough that a whole drag is one write, short enough that a kill moments
 * after the drag still leaves the settled geometry behind.
 */
export const WINDOW_BOUNDS_SAVE_DEBOUNCE_MS = 500

export interface WindowRect {
  x: number
  y: number
  width: number
  height: number
}

export interface ShellWindowState extends WindowRect {
  /** whether the window was maximized when its geometry was last saved */
  maximized: boolean
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

/**
 * The stored value as a state object, or null when it cannot be trusted.
 *
 * Fractional coordinates are rounded, not rejected: a hand-edited `x: 10.5`
 * still names a real position. Sizes below the window minimum or beyond any
 * display are rejected outright — they describe a window that cannot exist.
 */
export function parseShellWindowState(value: unknown): ShellWindowState | null {
  if (typeof value !== 'object' || value === null) return null
  const record = value as Record<string, unknown>
  if (
    !isFiniteNumber(record.x) ||
    !isFiniteNumber(record.y) ||
    !isFiniteNumber(record.width) ||
    !isFiniteNumber(record.height)
  ) {
    return null
  }
  const state: ShellWindowState = {
    x: Math.round(record.x),
    y: Math.round(record.y),
    width: Math.round(record.width),
    height: Math.round(record.height),
    maximized: record.maximized === true,
  }
  if (state.width < SHELL_WINDOW_MIN_WIDTH || state.height < SHELL_WINDOW_MIN_HEIGHT) return null
  if (state.width > MAX_DIMENSION || state.height > MAX_DIMENSION) return null
  return state
}

/**
 * Whether some display can still show the window's title strip — the part a
 * user grabs to bring the window back. The strip is the top of the window
 * (the tab bar doubles as the title bar on Windows/Linux); at least 120px of
 * it must overlap a work area, fully inside that area's height. A window
 * whose visible remainder is a sliver of its bottom edge fails here on
 * purpose: nothing left to grab means the user cannot move it back.
 */
export function isShellWindowVisible(state: WindowRect, workAreas: readonly WindowRect[]): boolean {
  const strip = Math.min(state.height, TITLE_STRIP_HEIGHT)
  return workAreas.some((area) => {
    const overlapWidth =
      Math.min(state.x + state.width, area.x + area.width) - Math.max(state.x, area.x)
    const overlapHeight =
      Math.min(state.y + strip, area.y + area.height) - Math.max(state.y, area.y)
    return overlapWidth >= MIN_VISIBLE_WIDTH && overlapHeight >= strip
  })
}

/**
 * The saved state when it exists, parses, and some display can still show it;
 * null means "open at the default placement".
 */
export function readShellWindowState(
  settingsPath: string,
  workAreas: readonly WindowRect[],
): ShellWindowState | null {
  const parsed = parseShellWindowState(readAppSettings(settingsPath)[WINDOW_BOUNDS_KEY])
  if (!parsed) return null
  return isShellWindowVisible(parsed, workAreas) ? parsed : null
}

/** The members the save path needs off a live window (BrowserWindow satisfies this). */
export interface ShellWindowSource {
  isMaximized(): boolean
  getNormalBounds(): WindowRect
}

/**
 * Write the window's normal-state geometry for the next launch.
 *
 * Fullscreen is deliberately not remembered: it reads as a transient mode for
 * presenting, not as a window arrangement to come back to.
 */
export function persistShellWindowState(settingsPath: string, win: ShellWindowSource): void {
  const { x, y, width, height } = win.getNormalBounds()
  writeAppSetting(settingsPath, WINDOW_BOUNDS_KEY, {
    x: Math.round(x),
    y: Math.round(y),
    width: Math.round(width),
    height: Math.round(height),
    maximized: win.isMaximized(),
  })
}

/**
 * Wire the save policy onto a live window: debounced while the user drags or
 * resizes — a crash or a kill then still leaves the last settled arrangement
 * behind — plus once on close, which is the moment the issue asks for and the
 * one that lands while a debounce is still pending. A vetoed close (the
 * unsaved-changes prompt) saves too; the geometry it is asking about has not
 * moved since, so the write is the same either way.
 */
export function installShellWindowBoundsPersistence(
  settingsPath: string,
  win: ShellWindowSource & {
    on(event: string, listener: () => void): unknown
    isDestroyed(): boolean
  },
): void {
  let saveTimer: ReturnType<typeof setTimeout> | null = null
  const save = (): void => {
    if (saveTimer) {
      clearTimeout(saveTimer)
      saveTimer = null
    }
    if (win.isDestroyed()) return
    try {
      persistShellWindowState(settingsPath, win)
    } catch (err) {
      // an unwritable settings file must not take the close (or the app) down
      console.warn('[shell] window bounds were not saved:', err)
    }
  }
  const schedule = (): void => {
    if (saveTimer) clearTimeout(saveTimer)
    saveTimer = setTimeout(save, WINDOW_BOUNDS_SAVE_DEBOUNCE_MS)
  }
  win.on('move', schedule)
  win.on('resize', schedule)
  win.on('close', save)
}
