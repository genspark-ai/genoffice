import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  installShellWindowBoundsPersistence,
  isShellWindowVisible,
  parseShellWindowState,
  persistShellWindowState,
  readShellWindowState,
  SHELL_WINDOW_MIN_HEIGHT,
  SHELL_WINDOW_MIN_WIDTH,
  WINDOW_BOUNDS_KEY,
  WINDOW_BOUNDS_SAVE_DEBOUNCE_MS,
} from '../src/main/window-bounds'

/**
 * The shell window's geometry across sessions (genoffice#2049): what is stored,
 * what is trusted back, and the debounced/close save policy.
 *
 * The settings file is a plain JSON file a user can hand-edit, so reading it
 * has to survive anything that ends up in there — that is most of these tests.
 * No Electron is needed: the save path takes a window-shaped object and the
 * module never imports `electron` itself.
 */

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'window-bounds-'))
  settingsPath = join(dir, 'app-settings.json')
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

function settingsFile(contents: unknown): string {
  writeFileSync(settingsPath, typeof contents === 'string' ? contents : JSON.stringify(contents))
  return settingsPath
}

describe('parseShellWindowState', () => {
  it('parses a well-formed saved state', () => {
    expect(
      parseShellWindowState({ x: 100, y: 80, width: 1200, height: 800, maximized: true }),
    ).toEqual({ x: 100, y: 80, width: 1200, height: 800, maximized: true })
  })

  it('rounds fractional coordinates instead of rejecting them', () => {
    expect(parseShellWindowState({ x: 10.4, y: 20.6, width: 1200.2, height: 800.7 })).toEqual({
      x: 10,
      y: 21,
      width: 1200,
      height: 801,
      maximized: false,
    })
  })

  it('defaults the maximized flag and ignores non-boolean values', () => {
    expect(parseShellWindowState({ x: 0, y: 0, width: 1200, height: 800 })?.maximized).toBe(false)
    expect(
      parseShellWindowState({ x: 0, y: 0, width: 1200, height: 800, maximized: 'yes' })?.maximized,
    ).toBe(false)
  })

  it('rejects anything that is not a rectangle of finite numbers', () => {
    expect(parseShellWindowState(null)).toBeNull()
    expect(parseShellWindowState(42)).toBeNull()
    expect(parseShellWindowState('1200x800')).toBeNull()
    expect(parseShellWindowState([])).toBeNull()
    expect(parseShellWindowState({})).toBeNull()
    expect(parseShellWindowState({ x: 0, y: 0, width: 1200 })).toBeNull()
    expect(parseShellWindowState({ x: 0, y: 0, width: NaN, height: 800 })).toBeNull()
    expect(parseShellWindowState({ x: 0, y: 0, width: Infinity, height: 800 })).toBeNull()
    expect(parseShellWindowState({ x: '100', y: 0, width: 1200, height: 800 })).toBeNull()
  })

  it('rejects sizes a real window cannot have', () => {
    const rect = { x: 0, y: 0, width: SHELL_WINDOW_MIN_WIDTH, height: SHELL_WINDOW_MIN_HEIGHT }
    expect(parseShellWindowState(rect)).not.toBeNull()
    expect(parseShellWindowState({ ...rect, width: SHELL_WINDOW_MIN_WIDTH - 1 })).toBeNull()
    expect(parseShellWindowState({ ...rect, height: SHELL_WINDOW_MIN_HEIGHT - 1 })).toBeNull()
    // the largest real displays are 8K — beyond that the value is corrupt
    expect(parseShellWindowState({ ...rect, width: 9000 })).toBeNull()
    expect(parseShellWindowState({ ...rect, height: 9000 })).toBeNull()
  })

  it('accepts negative coordinates (a display left of the primary)', () => {
    expect(parseShellWindowState({ x: -1400, y: 100, width: 1200, height: 800 })).not.toBeNull()
  })
})

describe('isShellWindowVisible', () => {
  const primary = { x: 0, y: 0, width: 1440, height: 900 }
  const secondary = { x: 1440, y: 0, width: 1920, height: 1080 }

  it('keeps a window inside the work area', () => {
    expect(isShellWindowVisible({ x: 100, y: 80, width: 1200, height: 800 }, [primary])).toBe(true)
  })

  it('keeps a window intentionally half off an edge, strip still showing', () => {
    // 240px of the window (and the whole title strip) stay on the display
    expect(isShellWindowVisible({ x: 1200, y: 80, width: 1360, height: 800 }, [primary])).toBe(true)
  })

  it('rejects a window saved on a monitor that is gone', () => {
    expect(isShellWindowVisible({ x: 2600, y: 100, width: 1200, height: 800 }, [primary])).toBe(
      false,
    )
  })

  it('accepts a window flush with the top, a few strip rows under the menu bar', () => {
    const belowMenuBar = { ...primary, y: 25, height: primary.height - 25 }
    expect(isShellWindowVisible({ x: 100, y: 0, width: 1200, height: 800 }, [belowMenuBar])).toBe(
      true,
    )
  })

  it('rejects a window pushed off the top — the strip is not reachable', () => {
    expect(isShellWindowVisible({ x: 100, y: -800, width: 1200, height: 800 }, [primary])).toBe(
      false,
    )
  })

  it('rejects a remaining sliver narrower than a grab', () => {
    // only 40px of the window's left edge shows
    expect(isShellWindowVisible({ x: 1400, y: 80, width: 1360, height: 800 }, [primary])).toBe(
      false,
    )
  })

  it('accepts a position on any of several displays', () => {
    expect(
      isShellWindowVisible({ x: 1600, y: 100, width: 1200, height: 800 }, [primary, secondary]),
    ).toBe(true)
  })
})

describe('readShellWindowState', () => {
  const primary = { x: 0, y: 0, width: 1440, height: 900 }

  it('returns null when the file or the key is missing', () => {
    expect(readShellWindowState(settingsPath, [primary])).toBeNull()
    settingsFile({ language: 'en' })
    expect(readShellWindowState(settingsPath, [primary])).toBeNull()
    settingsFile('{broken')
    expect(readShellWindowState(settingsPath, [primary])).toBeNull()
  })

  it('returns the saved state when a display can show it', () => {
    settingsFile({
      [WINDOW_BOUNDS_KEY]: { x: 120, y: 90, width: 960, height: 640, maximized: true },
    })
    expect(readShellWindowState(settingsPath, [primary])).toEqual({
      x: 120,
      y: 90,
      width: 960,
      height: 640,
      maximized: true,
    })
  })

  it('returns null when the saved position is off every display', () => {
    settingsFile({ [WINDOW_BOUNDS_KEY]: { x: 9000, y: 9000, width: 960, height: 640 } })
    expect(readShellWindowState(settingsPath, [primary])).toBeNull()
  })
})

describe('persistShellWindowState', () => {
  const fakeWindow = (
    bounds: { x: number; y: number; width: number; height: number },
    maximized = false,
  ) => ({
    getNormalBounds: () => bounds,
    isMaximized: () => maximized,
  })

  it('writes the normal-state rectangle with the maximized flag', () => {
    persistShellWindowState(settingsPath, fakeWindow({ x: 50, y: 60, width: 1000, height: 700 }))
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      [WINDOW_BOUNDS_KEY]: { x: 50, y: 60, width: 1000, height: 700, maximized: false },
    })
  })

  it('stores the normal bounds of a maximized window, not the maximized rectangle', () => {
    // the fake reports the pre-maximize rectangle as the normal bounds, like
    // Electron's getNormalBounds does for a maximized window
    persistShellWindowState(
      settingsPath,
      fakeWindow({ x: 50, y: 60, width: 1000, height: 700 }, true),
    )
    const saved = JSON.parse(readFileSync(settingsPath, 'utf8'))[WINDOW_BOUNDS_KEY]
    expect(saved).toEqual({ x: 50, y: 60, width: 1000, height: 700, maximized: true })
    expect(readShellWindowState(settingsPath, [{ x: 0, y: 0, width: 1440, height: 900 }])).toEqual(
      saved,
    )
  })

  it('keeps unrelated settings keys', () => {
    settingsFile({ language: 'ja' })
    persistShellWindowState(settingsPath, fakeWindow({ x: 0, y: 0, width: 1200, height: 800 }))
    expect(JSON.parse(readFileSync(settingsPath, 'utf8')).language).toBe('ja')
  })
})

interface FakeWindow {
  destroyed: boolean
  bounds: { x: number; y: number; width: number; height: number }
  getNormalBounds: () => { x: number; y: number; width: number; height: number }
  isMaximized: () => boolean
  isDestroyed: () => boolean
  on: (event: string, listener: () => void) => void
  emit: (event: string) => void
}

function makeFakeWindow(): FakeWindow {
  const handlers = new Map<string, Array<() => void>>()
  const win: FakeWindow = {
    destroyed: false,
    bounds: { x: 40, y: 50, width: 1000, height: 700 },
    getNormalBounds: vi.fn(() => ({ ...win.bounds })),
    isMaximized: () => false,
    isDestroyed: () => win.destroyed,
    on: (event, listener) => {
      const list = handlers.get(event) ?? []
      list.push(listener)
      handlers.set(event, list)
    },
    emit: (event) => {
      for (const listener of handlers.get(event) ?? []) listener()
    },
  }
  return win
}

describe('installShellWindowBoundsPersistence', () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })

  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('debounces a burst of moves and resizes into one save', () => {
    const win = makeFakeWindow()
    installShellWindowBoundsPersistence(settingsPath, win)
    win.emit('move')
    win.emit('resize')
    win.emit('move')
    expect(win.getNormalBounds).not.toHaveBeenCalled()

    vi.advanceTimersByTime(WINDOW_BOUNDS_SAVE_DEBOUNCE_MS - 1)
    expect(win.getNormalBounds).not.toHaveBeenCalled()
    vi.advanceTimersByTime(1)
    expect(win.getNormalBounds).toHaveBeenCalledTimes(1)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))[WINDOW_BOUNDS_KEY]).toMatchObject({
      x: 40,
      y: 50,
    })
  })

  it('saves immediately on close and drops a pending debounce', () => {
    const win = makeFakeWindow()
    installShellWindowBoundsPersistence(settingsPath, win)
    win.emit('move')
    win.emit('close')
    expect(win.getNormalBounds).toHaveBeenCalledTimes(1)

    vi.advanceTimersByTime(WINDOW_BOUNDS_SAVE_DEBOUNCE_MS * 4)
    expect(win.getNormalBounds).toHaveBeenCalledTimes(1)
    expect(readShellWindowState(settingsPath, [{ x: 0, y: 0, width: 1440, height: 900 }])).toEqual({
      x: 40,
      y: 50,
      width: 1000,
      height: 700,
      maximized: false,
    })
  })

  it('skips a destroyed window', () => {
    const win = makeFakeWindow()
    installShellWindowBoundsPersistence(settingsPath, win)
    win.destroyed = true
    expect(() => win.emit('close')).not.toThrow()
    expect(win.getNormalBounds).not.toHaveBeenCalled()
  })

  it('warns instead of throwing when the settings file is unwritable', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const blocked = join(dir, 'blocked')
    mkdirSync(blocked)
    const win = makeFakeWindow()
    installShellWindowBoundsPersistence(blocked, win)
    expect(() => win.emit('close')).not.toThrow()
    expect(warn).toHaveBeenCalled()
  })
})
