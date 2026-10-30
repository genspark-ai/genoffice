/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { LANGS, type Lang } from '@genoffice/i18n'
import type { DesktopPrefs, FolderRoot, HomeApi, RecentEntry } from '../src/shared/home-api'
import {
  DESKTOP_CELL_H,
  DESKTOP_CELL_W,
  DESKTOP_PAD_X,
  DESKTOP_PAD_Y,
  DesktopView,
} from '../src/renderer/src/DesktopView'
import { DESKTOP_STRINGS } from '../src/renderer/src/desktop-strings'

/**
 * The Home desktop against a mocked bridge: the icon set it renders, the
 * open paths of every icon kind, drag-to-rearrange with grid snap and the
 * debounced write-through, the context menu, and how unsaved icons flow into
 * free cells around saved ones.
 */

const DEBOUNCE_MS = 300

const actEnvironment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root
/** the layout the bridge hands back on load */
let stored: { positions: Record<string, { x: number; y: number }> }
/** every save, in order, with the exact map it carried */
let saves: Array<{ positions: Record<string, { x: number; y: number }> }>
let openPath: ReturnType<typeof vi.fn>
let revealPath: ReturnType<typeof vi.fn>
let newDoc: ReturnType<typeof vi.fn>
let newSheet: ReturnType<typeof vi.fn>
let newSlide: ReturnType<typeof vi.fn>
let newPdf: ReturnType<typeof vi.fn>
let newMarkdown: ReturnType<typeof vi.fn>
let newHtml: ReturnType<typeof vi.fn>
/** the recents page the bridge answers with */
let recentFiles: RecentEntry[]
/** the recents-changed handler the component subscribed */
let recentsChanged: (() => void) | undefined
/** the prefs the bridge hands back on load */
let storedPrefs: DesktopPrefs
/** every prefs save, in order, with the exact object it carried */
let prefSaves: DesktopPrefs[]
let statPaths: ReturnType<typeof vi.fn>
let getDefaultSaveDir: ReturnType<typeof vi.fn>
let wallpaperPickCustom: ReturnType<typeof vi.fn>
/** distinct data URL per wallpaper id, the way the main process serves them */
let wallpaperImage: ReturnType<typeof vi.fn>
/** the desktop-prefs-changed handler the component subscribed */
let prefsChanged: ((prefs: DesktopPrefs) => void) | undefined
const wallpaperUrl = (id: string): string => `data:image/png;base64,${btoa(id)}`

const entry = (path: string, name: string, ext: string): RecentEntry => ({
  path,
  name,
  ext,
  mtimeMs: 1,
  sizeBytes: 10,
  starred: false,
})

const fileA = entry('/docs/alpha.docx', 'alpha.docx', 'docx')
const fileB = entry('/docs/beta.xlsx', 'beta.xlsx', 'xlsx')
const fileC = entry('/docs/gamma.pdf', 'gamma.pdf', 'pdf')

const rootFolder: FolderRoot = {
  path: '/docs/Work',
  name: 'Work',
  usable: true,
  readable: true,
  removable: true,
}
const extraFolder: FolderRoot = {
  path: '/docs/Archive',
  name: 'Archive',
  usable: true,
  readable: true,
  removable: true,
}

const props = {
  lang: 'en' as Lang,
  roots: [rootFolder, extraFolder] as FolderRoot[],
  onOpenAsk: vi.fn(),
  onOpenTemplates: vi.fn(),
  onOpenSettings: vi.fn(),
  onOpenFolder: vi.fn(),
}

beforeEach(() => {
  vi.useFakeTimers()
  stored = { positions: {} }
  saves = []
  recentFiles = [fileA, fileB, fileC]
  recentsChanged = undefined
  prefsChanged = undefined
  storedPrefs = {
    wallpaper: 'default',
    carouselMinutes: 0,
    iconSize: 'md',
    sortMode: 'custom',
    wallpaperFit: 'cover',
    labelColor: '',
  }
  prefSaves = []
  openPath = vi.fn(async () => undefined)
  revealPath = vi.fn(async () => undefined)
  newDoc = vi.fn(async () => undefined)
  newSheet = vi.fn(async () => undefined)
  newSlide = vi.fn(async () => undefined)
  newPdf = vi.fn(async () => undefined)
  newMarkdown = vi.fn(async () => undefined)
  newHtml = vi.fn(async () => undefined)
  statPaths = vi.fn(async (paths: string[]) =>
    paths.map((path) => ({
      path,
      name: 'alpha.docx',
      ext: 'docx',
      mtimeMs: 1,
      sizeBytes: 1536,
      starred: false,
    })),
  )
  getDefaultSaveDir = vi.fn(async () => '/docs/GenOffice')
  wallpaperImage = vi.fn(async (id: string) => wallpaperUrl(id))
  wallpaperPickCustom = vi.fn(async () => 'custom:1700000000-photo.png')
  props.onOpenAsk.mockClear()
  props.onOpenTemplates.mockClear()
  props.onOpenSettings.mockClear()
  props.onOpenFolder.mockClear()
  // the surface measures itself through getBoundingClientRect: give every
  // element a 1200x800 rect at the viewport origin
  vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockReturnValue({
    width: 1200,
    height: 800,
    left: 0,
    top: 0,
    right: 1200,
    bottom: 800,
    x: 0,
    y: 0,
    toJSON: () => ({}),
  } as DOMRect)
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  window.aiOffice = {
    desktopLayout: vi.fn(async () => stored),
    desktopLayoutSave: vi.fn(
      async (layout: { positions: Record<string, { x: number; y: number }> }) => {
        saves.push(layout)
        return undefined
      },
    ),
    desktopPrefs: vi.fn(async () => storedPrefs),
    desktopPrefsSave: vi.fn(async (prefs: DesktopPrefs) => {
      prefSaves.push(prefs)
      return undefined
    }),
    recents: vi.fn(async (query: { limit?: number }) => ({
      entries: recentFiles.slice(0, query.limit ?? recentFiles.length),
      total: recentFiles.length,
      totalAll: recentFiles.length,
    })),
    openPath,
    revealPath,
    newDoc,
    newSheet,
    newSlide,
    newPdf,
    newMarkdown,
    newHtml,
    statPaths,
    getDefaultSaveDir,
    wallpaperImage,
    wallpaperPickCustom,
    onRecentsChanged: (handler: () => void) => {
      recentsChanged = handler
      return () => {}
    },
    onDesktopPrefsChanged: (handler: (prefs: DesktopPrefs) => void) => {
      prefsChanged = handler
      return () => {}
    },
  } as unknown as HomeApi
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  window.aiOffice = undefined as unknown as HomeApi
  vi.restoreAllMocks()
  vi.useRealTimers()
})

async function render(lang: Lang = 'en'): Promise<void> {
  await act(async () => {
    root.render(createElement(DesktopView, { ...props, lang }))
    await Promise.resolve()
  })
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

/** run the write-through debounce out */
async function tick(ms = DEBOUNCE_MS): Promise<void> {
  await act(async () => {
    await vi.advanceTimersByTimeAsync(ms)
  })
}

/** the pixel origin of grid cell (col, row) */
const cellPos = (col: number, row: number) => ({
  x: DESKTOP_PAD_X + col * DESKTOP_CELL_W,
  y: DESKTOP_PAD_Y + row * DESKTOP_CELL_H,
})

const icons = () => [...host.querySelectorAll<HTMLElement>('.desktop-icon')]
const labels = () => [...host.querySelectorAll('.desktop-icon-label')].map((l) => l.textContent)
const iconOf = (label: string): HTMLElement =>
  host.querySelector<HTMLElement>(`.desktop-icon-item[aria-label="${label}"]`)!
const posOf = (label: string): { x: string; y: string } => {
  const li = iconOf(label).closest('li') as HTMLElement
  return { x: li.style.left, y: li.style.top }
}
const posXY = (label: string): { x: number; y: number } => {
  const pos = posOf(label)
  return { x: Number.parseInt(pos.x, 10), y: Number.parseInt(pos.y, 10) }
}

/** a press-move-release gesture over the surface (origin at the viewport) */
async function dragTo(label: string, x: number, y: number): Promise<void> {
  const before = posXY(label)
  await act(async () => {
    iconOf(label).dispatchEvent(
      new MouseEvent('pointerdown', {
        bubbles: true,
        clientX: before.x,
        clientY: before.y,
        button: 0,
      }),
    )
    await Promise.resolve()
  })
  await act(async () => {
    window.dispatchEvent(new MouseEvent('pointermove', { bubbles: true, clientX: x, clientY: y }))
    await Promise.resolve()
  })
  await act(async () => {
    window.dispatchEvent(new MouseEvent('pointerup', { bubbles: true, clientX: x, clientY: y }))
    await Promise.resolve()
  })
}

async function click(el: Element | null | undefined): Promise<void> {
  expect(el).toBeTruthy()
  await act(async () => {
    ;(el as HTMLElement).click()
    await Promise.resolve()
  })
}

async function doubleClick(label: string): Promise<void> {
  await act(async () => {
    iconOf(label).dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
    await Promise.resolve()
  })
}

async function contextMenu(label: string): Promise<void> {
  await act(async () => {
    iconOf(label).dispatchEvent(
      new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 40 }),
    )
    await Promise.resolve()
  })
}

async function pressKey(key: string): Promise<void> {
  await act(async () => {
    window.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }))
    await Promise.resolve()
  })
}

const menuButtons = () =>
  [...host.querySelectorAll<HTMLButtonElement>('.folder-menu button')].map((b) => b.textContent)

/** run the pending microtasks so bridge answers land and re-render */
async function flushAsync(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
    await Promise.resolve()
  })
}

/** right-click the empty surface: the OS-grade menu with its submenus */
async function surfaceContextMenu(): Promise<void> {
  await act(async () => {
    host.querySelector('.desktop-surface')!.dispatchEvent(
      new MouseEvent('contextmenu', {
        bubbles: true,
        cancelable: true,
        clientX: 300,
        clientY: 300,
      }),
    )
    await Promise.resolve()
  })
}

/** open the wallpaper center off the surface menu, thumbnails fetched */
async function openWallpaperCenter(): Promise<void> {
  await surfaceContextMenu()
  await clickSurfaceButton('Change wallpaper')
  await flushAsync()
}

/** type into an input the way a user would (native setter + input event) */
async function typeInto(input: HTMLInputElement, text: string): Promise<void> {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
  await act(async () => {
    setter.call(input, text)
    input.dispatchEvent(new Event('input', { bubbles: true }))
    await Promise.resolve()
  })
}

/** the surface menu's own rows: the three submenu triggers plus the actions */
const surfaceMenuButtons = () =>
  [
    ...host.querySelectorAll<HTMLButtonElement>(
      '.desktop-menu > button, .desktop-menu > .desktop-menu-row > button',
    ),
  ].map((b) => b.textContent)

const submenuButtons = () =>
  [...host.querySelectorAll<HTMLButtonElement>('.desktop-submenu button')].map((b) => b.textContent)

/** click a surface-menu row by its exact label (trigger or submenu item) */
async function clickSurfaceButton(label: string): Promise<void> {
  const el = [...host.querySelectorAll<HTMLButtonElement>('.desktop-menu button')].find(
    (b) => b.textContent === label,
  )
  await click(el)
}

/** click a per-icon menu item by its exact label */
async function clickIconMenuButton(label: string): Promise<void> {
  const el = [
    ...host.querySelectorAll<HTMLButtonElement>('.folder-menu:not(.desktop-menu) button'),
  ].find((b) => b.textContent === label)
  await click(el)
}

/** open one of the surface menu's submenus by its trigger label */
async function openSurfaceSubmenu(trigger: string): Promise<void> {
  await clickSurfaceButton(trigger)
  expect(host.querySelector('.desktop-submenu')).toBeTruthy()
}

/** the file icons' labels, in the order the view lays them out */
const fileLabels = () => labels().filter((label) => label.includes('.'))

describe('DesktopView — the icon set', () => {
  it('renders the system apps, folder roots and newest recents as icons', async () => {
    await render()
    expect(icons()).toHaveLength(15)
    expect(labels()).toEqual([
      'Docs',
      'Sheets',
      'Slides',
      'PDF',
      'Markdown',
      'Web Page',
      'AI Chat',
      'Templates',
      'Plugins',
      'Settings',
      'Work',
      'Archive',
      'alpha.docx',
      'beta.xlsx',
      'gamma.pdf',
    ])
    expect(host.querySelector('.desktop-tile-docs')).toBeTruthy()
    expect(host.querySelector('.desktop-tile-ask')).toBeTruthy()
    // a file icon draws its FileBadge image; a folder a glass tile
    expect(iconOf('alpha.docx').querySelector('img')).toBeTruthy()
    expect(iconOf('Work').querySelector('.desktop-tile-folder')).toBeTruthy()
  })

  it('lays the icons out column-major from the top-left on a 96/104 cell grid', async () => {
    await render()
    expect(posOf('Docs')).toEqual({ x: '12px', y: '12px' })
    expect(posOf('Sheets')).toEqual({ x: `${cellPos(0, 1).x}px`, y: `${cellPos(0, 1).y}px` })
    expect(posOf('AI Chat')).toEqual({ x: `${cellPos(0, 6).x}px`, y: `${cellPos(0, 6).y}px` })
    expect(posOf('Templates')).toEqual({ x: `${cellPos(1, 0).x}px`, y: `${cellPos(1, 0).y}px` })
    expect(posOf('gamma.pdf')).toEqual({ x: `${cellPos(2, 0).x}px`, y: `${cellPos(2, 0).y}px` })
    // every icon owns a distinct cell
    const spots = icons().map((li) => `${li.style.left}:${li.style.top}`)
    expect(new Set(spots).size).toBe(spots.length)
  })

  it('shows at most the eight newest recents and skips missing files', async () => {
    recentFiles = [
      entry('/a/1.md', '1.md', 'md'),
      { ...entry('/a/2.md', '2.md', 'md'), missing: true },
      ...Array.from({ length: 9 }, (_v, i) => entry(`/a/f${i}.docx`, `f${i}.docx`, 'docx')),
    ]
    await render()
    const names = labels()
    expect(names).not.toContain('2.md')
    // the bridge paged eight; the missing one of those drops out client-side
    expect(names.filter((n) => n.startsWith('f') || n === '1.md')).toHaveLength(7)
    expect(icons()).toHaveLength(10 + 2 + 7)
  })

  it('survives a preload that predates the feature', async () => {
    window.aiOffice = {} as unknown as HomeApi
    await render()
    expect(icons()).toHaveLength(12) // the ten system apps plus the folder roots
    expect(saves).toEqual([])
  })
})

describe('DesktopView — opening icons', () => {
  it('double-click opens a recent file through openPath', async () => {
    await render()
    await doubleClick('alpha.docx')
    expect(openPath).toHaveBeenCalledWith('/docs/alpha.docx')
  })

  it('double-click on each system app calls its create action or parent callback', async () => {
    await render()
    await doubleClick('Docs')
    expect(newDoc).toHaveBeenCalledTimes(1)
    await doubleClick('AI Chat')
    expect(props.onOpenAsk).toHaveBeenCalledTimes(1)
    await doubleClick('Templates')
    expect(props.onOpenTemplates).toHaveBeenCalledTimes(1)
    await doubleClick('Plugins')
    expect(props.onOpenSettings).toHaveBeenCalledWith({ section: 'plugins' })
    await doubleClick('Settings')
    expect(props.onOpenSettings).toHaveBeenCalledWith({ section: 'general' })
  })

  it('double-click on a folder root opens that folder through the parent', async () => {
    await render()
    await doubleClick('Work')
    expect(props.onOpenFolder).toHaveBeenCalledWith(rootFolder)
  })

  it('Enter opens the selection; a single click only selects it', async () => {
    await render()
    await act(async () => {
      iconOf('beta.xlsx').dispatchEvent(
        new MouseEvent('pointerdown', { bubbles: true, clientX: 60, clientY: 40, button: 0 }),
      )
      await Promise.resolve()
    })
    await act(async () => {
      window.dispatchEvent(new MouseEvent('pointerup', { bubbles: true, clientX: 60, clientY: 40 }))
      await Promise.resolve()
    })
    expect(openPath).not.toHaveBeenCalled()
    expect(iconOf('beta.xlsx').closest('li')!.className).toContain('selected')
    await pressKey('Enter')
    expect(openPath).toHaveBeenCalledWith('/docs/beta.xlsx')
  })

  it('Enter does not reach the desktop while another surface owns the keyboard', async () => {
    await render()
    await act(async () => {
      iconOf('beta.xlsx').dispatchEvent(
        new MouseEvent('pointerdown', { bubbles: true, clientX: 60, clientY: 40, button: 0 }),
      )
      await Promise.resolve()
    })
    const outside = document.createElement('input')
    document.body.append(outside)
    await act(async () => {
      outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
      await Promise.resolve()
    })
    expect(openPath).not.toHaveBeenCalled()
    outside.remove()
  })

  it('the icon list is a proper listbox: options carry aria-selected', async () => {
    await render()
    const listbox = host.querySelector('[role="listbox"]')
    expect(listbox).toBeTruthy()
    const options = [...host.querySelectorAll<HTMLElement>('[role="option"]')]
    expect(options.length).toBeGreaterThan(0)
    expect(options.every((o) => o.getAttribute('aria-selected') === 'false')).toBe(true)
    await act(async () => {
      iconOf('beta.xlsx').dispatchEvent(
        new MouseEvent('pointerdown', { bubbles: true, clientX: 60, clientY: 40, button: 0 }),
      )
      await Promise.resolve()
    })
    await act(async () => {
      window.dispatchEvent(new MouseEvent('pointerup', { bubbles: true, clientX: 60, clientY: 40 }))
      await Promise.resolve()
    })
    expect(iconOf('beta.xlsx').getAttribute('aria-selected')).toBe('true')
  })

  it('arrow keys move the selection through the flow order', async () => {
    await render()
    await pressKey('ArrowDown')
    expect(host.querySelector('.desktop-icon.selected')?.textContent).toContain(
      host.querySelector('.desktop-icon')?.textContent ?? '',
    )
    // a second ArrowDown moves to the next icon, ArrowUp returns
    const first = host.querySelector('.desktop-icon.selected')?.textContent
    await pressKey('ArrowDown')
    const second = host.querySelector('.desktop-icon.selected')?.textContent
    expect(second).not.toBe(first)
    await pressKey('ArrowUp')
    expect(host.querySelector('.desktop-icon.selected')?.textContent).toBe(first)
  })

  it('type-ahead jumps to the first icon whose label starts with the typed run', async () => {
    await render()
    await pressKey('b')
    await pressKey('e')
    await pressKey('t')
    expect(host.querySelector('.desktop-icon.selected')?.textContent).toContain('beta.xlsx')
  })

  it('type-ahead stays silent while an input owns the keyboard', async () => {
    await render()
    const outside = document.createElement('input')
    document.body.append(outside)
    await act(async () => {
      outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'b', bubbles: true }))
      await Promise.resolve()
    })
    expect(host.querySelector('.desktop-icon.selected')).toBeNull()
    outside.remove()
  })

  it('Escape closes the menu first, then the selection', async () => {
    await render()
    await act(async () => {
      iconOf('alpha.docx').dispatchEvent(
        new MouseEvent('pointerdown', { bubbles: true, clientX: 12, clientY: 116, button: 0 }),
      )
      await Promise.resolve()
    })
    expect(iconOf('alpha.docx').closest('li')!.className).toContain('selected')
    await pressKey('Escape')
    expect(iconOf('alpha.docx').closest('li')!.className).not.toContain('selected')
    await contextMenu('alpha.docx')
    expect(host.querySelector('.folder-menu')).toBeTruthy()
    await pressKey('Escape')
    expect(host.querySelector('.folder-menu')).toBeNull()
  })

  it('a press on the empty backdrop clears the selection', async () => {
    await render()
    await act(async () => {
      iconOf('alpha.docx').dispatchEvent(
        new MouseEvent('pointerdown', { bubbles: true, clientX: 12, clientY: 116, button: 0 }),
      )
      await Promise.resolve()
    })
    expect(iconOf('alpha.docx').closest('li')!.className).toContain('selected')
    const surface = host.querySelector<HTMLElement>('.desktop-surface')!
    await act(async () => {
      surface.dispatchEvent(
        new MouseEvent('pointerdown', { bubbles: true, clientX: 700, clientY: 700, button: 0 }),
      )
      await Promise.resolve()
    })
    expect(iconOf('alpha.docx').closest('li')!.className).not.toContain('selected')
  })
})

describe('DesktopView — drag to rearrange', () => {
  it('a drag snaps to the grid cell and persists after the 300ms debounce', async () => {
    await render()
    // grab Docs by its origin (12,12) and drop it over cell (6,4)
    await dragTo('Docs', cellPos(6, 4).x, cellPos(6, 4).y)
    expect(posOf('Docs')).toEqual({ x: `${cellPos(6, 4).x}px`, y: `${cellPos(6, 4).y}px` })
    // the vacated cell is re-taken by the next free icon (auto flow)
    expect(posOf('Sheets')).toEqual({ x: '12px', y: '12px' })
    expect(saves).toEqual([])
    await tick(DEBOUNCE_MS - 1)
    expect(saves).toEqual([])
    await tick(1)
    expect(saves).toHaveLength(1)
    expect(saves[0]).toEqual({ positions: { 'app:docs': cellPos(6, 4) } })
  })

  it('a drag onto an occupied cell swaps the two icons', async () => {
    await render()
    // Docs sits at (0,0); drag it onto Sheets' cell (0,1)
    await dragTo('Docs', cellPos(0, 1).x, cellPos(0, 1).y)
    expect(posOf('Docs')).toEqual({ x: `${cellPos(0, 1).x}px`, y: `${cellPos(0, 1).y}px` })
    expect(posOf('Sheets')).toEqual({ x: '12px', y: '12px' })
    await tick()
    expect(saves[0]).toEqual({
      positions: { 'app:docs': cellPos(0, 1), 'app:sheets': cellPos(0, 0) },
    })
  })

  it("a drop onto the icon's own cell changes nothing and saves nothing", async () => {
    await render()
    await dragTo('Docs', cellPos(0, 0).x + 6, cellPos(0, 0).y + 6)
    expect(posOf('Docs')).toEqual({ x: '12px', y: '12px' })
    await tick()
    expect(saves).toEqual([])
  })

  it('a press that never crosses the drag threshold is a click, not a move', async () => {
    await render()
    await dragTo('Docs', cellPos(0, 0).x + 2, cellPos(0, 0).y + 2)
    expect(posOf('Docs')).toEqual({ x: '12px', y: '12px' })
    expect(saves).toEqual([])
  })

  it('a burst of drops is one save of the final map', async () => {
    await render()
    await dragTo('Docs', cellPos(4, 0).x, cellPos(4, 0).y)
    await dragTo('Docs', cellPos(5, 2).x, cellPos(5, 2).y)
    await tick()
    expect(saves).toHaveLength(1)
    expect(saves[0]).toEqual({ positions: { 'app:docs': cellPos(5, 2) } })
  })

  it('saved positions are restored on load, and unsaved icons flow around them', async () => {
    stored = {
      positions: {
        'app:docs': cellPos(3, 2),
        'file:/docs/gamma.pdf': cellPos(0, 0),
      },
    }
    await render()
    expect(posOf('Docs')).toEqual({ x: `${cellPos(3, 2).x}px`, y: `${cellPos(3, 2).y}px` })
    expect(posOf('gamma.pdf')).toEqual({ x: '12px', y: '12px' })
    // the unpinned icons flow into the remaining cells in reading order
    expect(posOf('Sheets')).toEqual({ x: `${cellPos(0, 1).x}px`, y: `${cellPos(0, 1).y}px` })
    // no icon overlaps a saved one
    const spots = icons().map((li) => `${li.style.left}:${li.style.top}`)
    expect(new Set(spots).size).toBe(spots.length)
    // and nothing is written back: the file already says this
    await tick()
    expect(saves).toEqual([])
  })
})

describe('DesktopView — icons appearing later', () => {
  it('a new recent file flows into a free cell without overlapping the saved ones', async () => {
    stored = { positions: { 'app:docs': cellPos(0, 0) } }
    await render()
    expect(posOf('Docs')).toEqual({ x: '12px', y: '12px' })
    expect(posOf('Sheets')).toEqual({ x: `${cellPos(0, 1).x}px`, y: `${cellPos(0, 1).y}px` })
    // a save lands in an editor tab: the recents list grows underneath us.
    // delta.md is the newest, so it enters the flow ahead of alpha and the
    // later siblings shift one cell — none of them overlapping the saved one
    const fresh = entry('/docs/delta.md', 'delta.md', 'md')
    recentFiles = [fresh, ...recentFiles]
    await act(async () => {
      recentsChanged!()
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(labels()).toContain('delta.md')
    expect(posOf('delta.md')).toEqual({ x: `${cellPos(1, 5).x}px`, y: `${cellPos(1, 5).y}px` })
    expect(posOf('gamma.pdf')).toEqual({ x: `${cellPos(2, 1).x}px`, y: `${cellPos(2, 1).y}px` })
    const spots = icons().map((li) => `${li.style.left}:${li.style.top}`)
    expect(new Set(spots).size).toBe(spots.length)
    expect(posOf('Docs')).toEqual({ x: '12px', y: '12px' })
    expect(saves).toEqual([]) // nothing the person did: nothing written
  })
})

describe('DesktopView — context menu', () => {
  it('offers open / reveal / properties / reset on a file icon and acts on each', async () => {
    await render()
    await contextMenu('alpha.docx')
    expect(menuButtons()).toEqual(['Open', 'Reveal in folder', 'Properties', 'Reset position'])
    await click([...host.querySelectorAll<HTMLButtonElement>('.folder-menu button')][0])
    expect(openPath).toHaveBeenCalledWith('/docs/alpha.docx')
    expect(host.querySelector('.folder-menu')).toBeNull()

    await contextMenu('alpha.docx')
    await click([...host.querySelectorAll<HTMLButtonElement>('.folder-menu button')][1])
    expect(revealPath).toHaveBeenCalledWith('/docs/alpha.docx')
  })

  it('offers no reveal on app icons, and properties only on file and folder icons', async () => {
    await render()
    await contextMenu('Docs')
    expect(menuButtons()).toEqual(['Open', 'Reset position'])
    await contextMenu('Work')
    expect(menuButtons()).toEqual(['Open', 'Properties', 'Reset position'])
  })

  it('an outside press closes the open menu', async () => {
    await render()
    await contextMenu('alpha.docx')
    expect(host.querySelector('.folder-menu')).toBeTruthy()
    await act(async () => {
      host
        .querySelector('.desktop-surface')!
        .dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientX: 700, clientY: 700 }))
      await Promise.resolve()
    })
    expect(host.querySelector('.folder-menu')).toBeNull()
  })

  it("reset drops the icon's saved position: it reflows and the save drops it", async () => {
    stored = { positions: { 'app:docs': cellPos(4, 4) } }
    await render()
    expect(posOf('Docs')).toEqual({ x: `${cellPos(4, 4).x}px`, y: `${cellPos(4, 4).y}px` })
    await contextMenu('Docs')
    await click([...host.querySelectorAll<HTMLButtonElement>('.folder-menu button')].at(-1))
    // the icon returns to the reading-order flow
    expect(posOf('Docs')).toEqual({ x: '12px', y: '12px' })
    await tick()
    expect(saves).toHaveLength(1)
    expect(saves[0]).toEqual({ positions: {} })
  })
})

describe('desktop strings — every UI language', () => {
  it('covers all 21 langs with non-empty words for each key', () => {
    expect(LANGS).toHaveLength(21)
    for (const lang of LANGS) {
      const words = DESKTOP_STRINGS[lang]
      expect(words, `${lang} is present`).toBeTruthy()
      for (const [key, value] of Object.entries(words)) {
        expect((value as string).length, `${lang}.${key} is empty`).toBeGreaterThan(0)
      }
    }
  })

  it('renders the zh words when the lang prop is zh', async () => {
    await render('zh')
    expect(labels()).toContain('文档')
    expect(labels()).toContain('AI 问答')
    expect(host.querySelector('.desktop-icons')!.getAttribute('aria-label')).toBe('桌面图标')
  })
})

describe('DesktopView — the empty-surface menu', () => {
  it('right-click on empty surface offers the OS-grade menu, submenus closed', async () => {
    await render()
    await surfaceContextMenu()
    expect(surfaceMenuButtons()).toEqual([
      'New',
      'Sort by',
      'View',
      'Refresh',
      'Change wallpaper',
      'Properties',
    ])
    expect(host.querySelector('.desktop-submenu')).toBeNull()
    // an icon's menu and the surface menu never coexist
    await contextMenu('alpha.docx')
    expect(host.querySelector('.desktop-menu')).toBeNull()
    expect(host.querySelector('.folder-menu')).toBeTruthy()
    await surfaceContextMenu()
    expect(host.querySelector('.folder-menu:not(.desktop-menu)')).toBeNull()
    expect(host.querySelector('.desktop-menu')).toBeTruthy()
  })

  it('an outside press closes the surface menu', async () => {
    await render()
    await surfaceContextMenu()
    expect(host.querySelector('.desktop-menu')).toBeTruthy()
    await act(async () => {
      host
        .querySelector('.desktop-surface')!
        .dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientX: 700, clientY: 700 }))
      await Promise.resolve()
    })
    expect(host.querySelector('.desktop-menu')).toBeNull()
  })

  it('新建 offers the six document types and runs the matching new* call', async () => {
    await render()
    await surfaceContextMenu()
    await openSurfaceSubmenu('New')
    expect(submenuButtons()).toEqual(['Docs', 'Sheets', 'Slides', 'PDF', 'Markdown', 'Web Page'])
    await clickSurfaceButton('Sheets')
    expect(newSheet).toHaveBeenCalledTimes(1)
    expect(host.querySelector('.desktop-menu')).toBeNull()
    // the menu re-opens for the next creation
    await surfaceContextMenu()
    await openSurfaceSubmenu('New')
    await clickSurfaceButton('PDF')
    expect(newPdf).toHaveBeenCalledTimes(1)
    expect(newDoc).not.toHaveBeenCalled()
  })

  it('刷新 re-pulls the recents', async () => {
    await render()
    const calls = (window.aiOffice.recents as ReturnType<typeof vi.fn>).mock.calls.length
    await surfaceContextMenu()
    await clickSurfaceButton('Refresh')
    expect((window.aiOffice.recents as ReturnType<typeof vi.fn>).mock.calls).toHaveLength(calls + 1)
    expect(host.querySelector('.desktop-menu')).toBeNull()
  })
})

describe('DesktopView — sort modes', () => {
  it('the 排序方式 submenu lists the four modes with the current one checked', async () => {
    await render()
    await surfaceContextMenu()
    await openSurfaceSubmenu('Sort by')
    expect(submenuButtons()).toEqual(['Custom (drag placement)', 'Name', 'Type', 'Date modified'])
    const checked = [...host.querySelectorAll<HTMLButtonElement>('.desktop-submenu button.checked')]
    expect(checked.map((b) => b.textContent)).toEqual(['Custom (drag placement)'])
  })

  it('name sort reorders the icons, disables dragging and persists the mode', async () => {
    recentFiles = [fileC, fileA, fileB] // gamma, alpha, beta — out of name order
    await render()
    expect(fileLabels()).toEqual(['gamma.pdf', 'alpha.docx', 'beta.xlsx'])
    await surfaceContextMenu()
    await openSurfaceSubmenu('Sort by')
    await clickSurfaceButton('Name')
    expect(fileLabels()).toEqual(['alpha.docx', 'beta.xlsx', 'gamma.pdf'])
    expect(prefSaves.at(-1)).toMatchObject({ sortMode: 'name' })
    // read-only flow: a drag gesture selects nothing and moves nothing
    // (Docs flows at cell (0,4) in name order: AI Chat, alpha, Archive, beta, Docs)
    const before = posOf('Docs')
    expect(before).toEqual({ x: '12px', y: `${cellPos(0, 4).y}px` })
    await dragTo('Docs', cellPos(6, 4).x, cellPos(6, 4).y)
    expect(posOf('Docs')).toEqual(before)
    expect(saves).toEqual([])
  })

  it('kind and mtime sorts order the files by type and newest-first', async () => {
    recentFiles = [
      { ...fileA, mtimeMs: 30 },
      { ...fileB, mtimeMs: 10 },
      { ...fileC, mtimeMs: 20 },
    ]
    await render()
    await surfaceContextMenu()
    await openSurfaceSubmenu('Sort by')
    await clickSurfaceButton('Type')
    // docx < pdf < xlsx at the file tier
    expect(labels().slice(-3)).toEqual(['alpha.docx', 'gamma.pdf', 'beta.xlsx'])
    await surfaceContextMenu()
    await openSurfaceSubmenu('Sort by')
    await clickSurfaceButton('Date modified')
    expect(labels().slice(-3)).toEqual(['alpha.docx', 'gamma.pdf', 'beta.xlsx'])
    expect(prefSaves.at(-1)).toMatchObject({ sortMode: 'mtime' })
  })

  it('choosing 自定义 again restores the saved free layout untouched', async () => {
    stored = { positions: { 'app:docs': cellPos(3, 2), 'file:/docs/alpha.docx': cellPos(0, 0) } }
    recentFiles = [fileC, fileA, fileB]
    await render()
    expect(posOf('Docs')).toEqual({ x: `${cellPos(3, 2).x}px`, y: `${cellPos(3, 2).y}px` })
    expect(posOf('alpha.docx')).toEqual({ x: '12px', y: '12px' })
    await surfaceContextMenu()
    await openSurfaceSubmenu('Sort by')
    await clickSurfaceButton('Name')
    // the read-only flow ignores the saved map: everything flows in name order
    // (AI Chat, alpha.docx, Archive, beta.xlsx, Docs, … on a 7-row grid)
    expect(posOf('Docs')).toEqual({ x: '12px', y: `${cellPos(0, 4).y}px` })
    expect(posOf('alpha.docx')).toEqual({ x: '12px', y: `${cellPos(0, 1).y}px` })
    await surfaceContextMenu()
    await openSurfaceSubmenu('Sort by')
    await clickSurfaceButton('Custom (drag placement)')
    // and back: the pinned cells are exactly where they were
    expect(posOf('Docs')).toEqual({ x: `${cellPos(3, 2).x}px`, y: `${cellPos(3, 2).y}px` })
    expect(posOf('alpha.docx')).toEqual({ x: '12px', y: '12px' })
    expect(prefSaves.at(-1)).toMatchObject({ sortMode: 'custom' })
    // the mode switch never wrote the layout: only the two prefs saves exist
    expect(saves).toEqual([])
  })

  it('loads the saved sort mode and icon size on launch, writing nothing back', async () => {
    storedPrefs = {
      wallpaper: 'default',
      carouselMinutes: 0,
      iconSize: 'lg',
      sortMode: 'kind',
      wallpaperFit: 'tile',
      labelColor: '#0f7fff',
    }
    recentFiles = [fileC, fileA, fileB]
    await render()
    expect(host.querySelector('.desktop-content')!.className).toContain('icons-lg')
    expect(labels().slice(-3)).toEqual(['alpha.docx', 'gamma.pdf', 'beta.xlsx'])
    expect(prefSaves).toEqual([])
  })
})

describe('DesktopView — icon size', () => {
  it('查看 ▸ 图标大小 applies the size class on the desktop root and persists', async () => {
    await render()
    expect(host.querySelector('.desktop-content')!.className).toContain('icons-md')
    await surfaceContextMenu()
    await openSurfaceSubmenu('View')
    expect(submenuButtons()).toEqual(['Small', 'Medium', 'Large'])
    await clickSurfaceButton('Large')
    expect(host.querySelector('.desktop-content')!.className).toContain('icons-lg')
    expect(prefSaves.at(-1)).toMatchObject({ iconSize: 'lg' })
    // the choice survives a reopen of the submenu
    await surfaceContextMenu()
    await openSurfaceSubmenu('View')
    const checked = [...host.querySelectorAll<HTMLButtonElement>('.desktop-submenu button.checked')]
    expect(checked.map((b) => b.textContent)).toEqual(['Large'])
    await clickSurfaceButton('Small')
    expect(host.querySelector('.desktop-content')!.className).toContain('icons-sm')
  })
})

describe('DesktopView — the icon properties dialog', () => {
  it('shows the statPaths data for a file icon and reveals its folder', async () => {
    await render()
    await contextMenu('alpha.docx')
    await clickIconMenuButton('Properties')
    expect(host.querySelector('.desktop-props-dialog')).toBeTruthy()
    expect(statPaths).toHaveBeenCalledWith(['/docs/alpha.docx'])
    await flushAsync()
    const rows = () =>
      [...host.querySelectorAll('.desktop-props-row')].map(
        (row) =>
          `${row.querySelector('dt')!.textContent}:${row.querySelector('dd')!.textContent ?? ''}`,
      )
    expect(rows().slice(0, 3)).toEqual(['Name:alpha.docx', 'Type:DOCX', 'Size:1.5 KB'])
    // localized for the UI language; the epoch date renders in the machine's
    // own timezone (UTC or east of it lands on 1/1/1970, west on 12/31/1969)
    expect(rows()[3]).toMatch(/^Modified:(12\/31\/1969|1\/1\/1970), /)
    expect(rows()[4]).toBe('Path:')
    const pathInput = host.querySelector<HTMLInputElement>('.desktop-props-path')!
    expect(pathInput.value).toBe('/docs/alpha.docx')
    expect(pathInput.readOnly).toBe(true)
    await click(host.querySelector('.desktop-props-reveal'))
    expect(revealPath).toHaveBeenCalledWith('/docs/alpha.docx')
    // Escape closes the dialog before it ever touches the selection
    await pressKey('Escape')
    expect(host.querySelector('.desktop-props-dialog')).toBeNull()
  })

  it('shows the folder type without size or modified rows on a folder icon', async () => {
    await render()
    await contextMenu('Work')
    await clickIconMenuButton('Properties')
    expect(statPaths).not.toHaveBeenCalled()
    await flushAsync()
    const rows = [...host.querySelectorAll('.desktop-props-row')].map(
      (row) =>
        `${row.querySelector('dt')!.textContent}:${row.querySelector('dd')!.textContent ?? ''}`,
    )
    expect(rows).toEqual(['Name:Work', 'Type:Folder', 'Path:'])
    expect(host.querySelector<HTMLInputElement>('.desktop-props-path')!.value).toBe('/docs/Work')
  })

  it('an outside press closes the dialog', async () => {
    await render()
    await contextMenu('alpha.docx')
    await clickIconMenuButton('Properties')
    expect(host.querySelector('.desktop-props-dialog')).toBeTruthy()
    await act(async () => {
      host
        .querySelector('.set-overlay')!
        .dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 5, clientY: 5 }))
      await Promise.resolve()
    })
    expect(host.querySelector('.desktop-props-dialog')).toBeNull()
  })
})

describe('DesktopView — the desktop properties dialog', () => {
  it('shows the file and folder counts, the save dir, and opens the settings center', async () => {
    await render()
    await surfaceContextMenu()
    await clickSurfaceButton('Properties')
    expect(host.querySelector('.desktop-props-dialog')).toBeTruthy()
    await flushAsync()
    expect(getDefaultSaveDir).toHaveBeenCalledTimes(1)
    const rows = [...host.querySelectorAll('.desktop-props-row')].map(
      (row) =>
        `${row.querySelector('dt')!.textContent}:${row.querySelector('dd')!.textContent ?? ''}`,
    )
    expect(rows).toEqual(['Files:3', 'Folders:2', 'Default save location:'])
    expect(host.querySelector<HTMLInputElement>('.desktop-props-path')!.value).toBe(
      '/docs/GenOffice',
    )
    await click(host.querySelector('.desktop-props-open-settings'))
    expect(props.onOpenSettings).toHaveBeenCalledWith({ section: 'general' })
    expect(host.querySelector('.desktop-props-dialog')).toBeNull()
  })
})

describe('DesktopView — prefs changed outside the renderer', () => {
  /** the backdrop's active crossfade layer */
  const activeLayer = (): HTMLElement | null =>
    host.querySelector<HTMLElement>('.desktop-wallpaper > div.active')

  it('a desktop-prefs-changed push re-pulls the prefs and swaps the wallpaper live', async () => {
    await render()
    expect(host.querySelector('.desktop-aurora-a')).toBeTruthy() // the default aurora
    const prefsCalls = (window.aiOffice.desktopPrefs as ReturnType<typeof vi.fn>).mock.calls.length
    // the CLI/MCP desktop domain wrote a new wallpaper and broadcast
    storedPrefs = { ...storedPrefs, wallpaper: 'builtin:teal-deep' }
    await act(async () => {
      prefsChanged!(storedPrefs)
      await Promise.resolve()
      await Promise.resolve()
    })
    await flushAsync()
    // the prefs were re-pulled (not trusted from the event payload) and applied
    expect((window.aiOffice.desktopPrefs as ReturnType<typeof vi.fn>).mock.calls).toHaveLength(
      prefsCalls + 1,
    )
    expect(activeLayer()?.style.backgroundImage).toContain(wallpaperUrl('builtin:teal-deep'))
    expect(host.querySelector('.desktop-aurora-a')).toBeNull()
    expect(wallpaperImage).toHaveBeenCalledWith('builtin:teal-deep')
    // a remote change is applied, never written back
    expect(prefSaves).toEqual([])
  })

  it('survives a preload without onDesktopPrefsChanged', async () => {
    window.aiOffice = {
      desktopLayout: vi.fn(async () => stored),
      desktopPrefs: vi.fn(async () => storedPrefs),
    } as unknown as HomeApi
    await render()
    expect(icons()).toHaveLength(12) // the ten system apps plus the folder roots
  })
})

describe('DesktopView — the wallpaper center', () => {
  /** the backdrop's crossfade layers as (backgroundImage, active) pairs */
  const layers = (): Array<{ image: string; active: boolean }> =>
    [...host.querySelectorAll<HTMLElement>('.desktop-wallpaper > div')].map((el) => ({
      image: el.style.backgroundImage,
      active: el.classList.contains('active'),
    }))
  const activeLayer = (): HTMLElement | null =>
    host.querySelector<HTMLElement>('.desktop-wallpaper > div.active')

  it('applies a builtin wallpaper, persists it and updates the backdrop', async () => {
    await render()
    // the default look: the CSS aurora, no image layer
    expect(host.querySelector('.desktop-aurora-a')).toBeTruthy()
    expect(host.querySelector('.desktop-wallpaper > div')).toBeNull()
    await surfaceContextMenu()
    await clickSurfaceButton('Change wallpaper')
    expect(host.querySelector('.desktop-wp-dialog')).toBeTruthy()
    await flushAsync()
    // live thumbnails: every static plus the animated starfield PNG
    expect(host.querySelectorAll('.desktop-wp-tile img').length).toBe(6)
    const tile = [...host.querySelectorAll<HTMLElement>('.desktop-wp-tile')].find((t) =>
      t.textContent?.includes('Minimal light'),
    )
    await click(tile)
    expect(prefSaves.at(-1)).toMatchObject({ wallpaper: 'builtin:minimal-light' })
    expect(host.querySelector('.desktop-wp-dialog')).toBeNull() // applied and closed
    await flushAsync()
    const layer = activeLayer()
    expect(layer?.style.backgroundImage).toContain(wallpaperUrl('builtin:minimal-light'))
    expect(host.querySelector('.desktop-aurora-a')).toBeNull()
    expect(wallpaperImage).toHaveBeenCalledWith('builtin:minimal-light')
  })

  it('applies the picked custom image through wallpaperPickCustom', async () => {
    await render()
    await surfaceContextMenu()
    await clickSurfaceButton('Change wallpaper')
    await flushAsync()
    await click(
      [...host.querySelectorAll<HTMLElement>('.desktop-wp-tile')].find((t) =>
        t.textContent?.includes('Custom image'),
      ),
    )
    expect(wallpaperPickCustom).toHaveBeenCalledTimes(1)
    await flushAsync()
    expect(prefSaves.at(-1)).toMatchObject({ wallpaper: 'custom:1700000000-photo.png' })
    expect(activeLayer()?.style.backgroundImage).toContain(
      wallpaperUrl('custom:1700000000-photo.png'),
    )
  })

  it('the carousel timer advances the slideshow and crossfades without transforms', async () => {
    await render()
    await surfaceContextMenu()
    await clickSurfaceButton('Change wallpaper')
    const minute = [
      ...host.querySelectorAll<HTMLButtonElement>('.desktop-wp-carousel button'),
    ].find((b) => b.textContent === '1 min')
    await click(minute)
    expect(prefSaves.at(-1)).toMatchObject({ carouselMinutes: 1 })
    await click(host.querySelector('.desktop-wp-dialog .set-close'))
    expect(host.querySelector('.desktop-wp-dialog')).toBeNull()
    await flushAsync()
    // the slideshow starts at the first static and owns the backdrop
    expect(activeLayer()?.style.backgroundImage).toContain(wallpaperUrl('builtin:brand-aurora'))
    await act(async () => {
      await vi.advanceTimersByTimeAsync(60_000)
    })
    await flushAsync()
    expect(activeLayer()?.style.backgroundImage).toContain(wallpaperUrl('builtin:warm-dusk'))
    // the crossfade keeps the previous image mounted, fading out under it
    expect(layers().map((l) => l.image)).toEqual(
      expect.arrayContaining([
        expect.stringContaining(wallpaperUrl('builtin:brand-aurora')),
        expect.stringContaining(wallpaperUrl('builtin:warm-dusk')),
      ]),
    )
    // a wallpaper choice mid-slideshow pins the start back on that static
    expect(prefSaves.filter((p) => p.carouselMinutes === 1)).toHaveLength(1)
  })
})

describe('DesktopView — the wallpaper fit', () => {
  const fitButtons = () => [...host.querySelectorAll<HTMLButtonElement>('.desktop-wp-fit button')]
  const fitBy = (label: string): HTMLButtonElement =>
    fitButtons().find((b) => b.textContent === label)!
  const layer = () => host.querySelector<HTMLElement>('.desktop-wallpaper > div.active')!

  it('hides the row for the animated aurora, shows it for an image wallpaper', async () => {
    await render()
    await openWallpaperCenter()
    // the CSS aurora has no image to fill
    expect(host.querySelector('.desktop-wp-fit')).toBeNull()
    await click(
      [...host.querySelectorAll<HTMLElement>('.desktop-wp-tile')].find((t) =>
        t.textContent?.includes('Minimal light'),
      ),
    )
    await openWallpaperCenter()
    expect(fitButtons().map((b) => b.textContent)).toEqual([
      'Fill',
      'Fit',
      'Stretch',
      'Tile',
      'Center',
    ])
    expect(fitButtons().find((b) => b.classList.contains('active'))?.textContent).toBe('Fill')
  })

  it('applies every fit to the backdrop layer — tile repeats from the origin', async () => {
    storedPrefs = { ...storedPrefs, wallpaper: 'builtin:teal-deep' }
    await render()
    await flushAsync()
    await openWallpaperCenter()
    // the shipped default
    expect(layer().style.backgroundSize).toBe('cover')
    expect(layer().style.backgroundRepeat).toBe('no-repeat')
    await click(fitBy('Tile'))
    expect(prefSaves.at(-1)).toMatchObject({ wallpaperFit: 'tile' })
    expect(layer().style.backgroundRepeat).toBe('repeat')
    expect(layer().style.backgroundSize).toBe('auto')
    expect(layer().style.backgroundPosition).toBe('left top')
    await click(fitBy('Center'))
    expect(layer().style.backgroundRepeat).toBe('no-repeat')
    expect(layer().style.backgroundSize).toBe('auto')
    // cssstyle expands the single keyword to both axes
    expect(layer().style.backgroundPosition).toContain('center')
    await click(fitBy('Stretch'))
    expect(layer().style.backgroundSize).toBe('100% 100%')
    await click(fitBy('Fit'))
    expect(layer().style.backgroundSize).toBe('contain')
    expect(layer().style.backgroundPosition).toContain('center')
  })

  it('the carousel crossfade layers all honor the chosen fit', async () => {
    storedPrefs = { ...storedPrefs, wallpaper: 'builtin:teal-deep', carouselMinutes: 1 }
    await render()
    await openWallpaperCenter()
    await click(fitBy('Tile'))
    await click(host.querySelector('.desktop-wp-dialog .set-close'))
    await flushAsync()
    await act(async () => {
      await vi.advanceTimersByTimeAsync(60_000)
    })
    await flushAsync()
    // the fading-out layer and the new one both carry the repeat
    const both = [...host.querySelectorAll<HTMLElement>('.desktop-wallpaper > div')]
    expect(both.length).toBe(2)
    for (const el of both) {
      expect(el.style.backgroundRepeat).toBe('repeat')
      expect(el.style.backgroundSize).toBe('auto')
    }
  })

  it('the selected thumbnail previews the fit; the others keep the neutral crop', async () => {
    storedPrefs = { ...storedPrefs, wallpaper: 'builtin:minimal-light' }
    await render()
    await openWallpaperCenter()
    const thumb = () => host.querySelector<HTMLImageElement>('.desktop-wp-tile.active img')
    expect(thumb()!.style.objectFit).toBe('cover')
    await click(fitBy('Tile'))
    expect(thumb()!.style.objectFit).toBe('none')
    expect(thumb()!.style.objectPosition).toBe('left top')
    const inactive = [...host.querySelectorAll<HTMLElement>('.desktop-wp-tile')].find(
      (t) => !t.classList.contains('active') && t.querySelector('img'),
    )!
    expect(inactive.querySelector('img')!.style.objectFit).toBe('')
  })
})

describe('DesktopView — the label color', () => {
  const swatchBy = (label: string): HTMLButtonElement =>
    [...host.querySelectorAll<HTMLButtonElement>('.desktop-wp-swatch')].find(
      (b) => b.getAttribute('aria-label') === label,
    )!
  const hex = () => host.querySelector<HTMLInputElement>('.desktop-wp-hex')!
  /** the label-color custom properties the view sets on its root */
  const labelVar = (name: string): string =>
    host.querySelector<HTMLElement>('.desktop-content')!.style.getPropertyValue(name)

  it('offers the automatic reset plus the white/dark/brand presets', async () => {
    await render()
    await openWallpaperCenter()
    const swatches = [...host.querySelectorAll<HTMLButtonElement>('.desktop-wp-swatch')]
    expect(swatches.map((b) => b.getAttribute('aria-label'))).toEqual([
      'Auto',
      'White',
      'Dark gray',
      'Brand blue',
      'Brand green',
      'Brand orange',
      'Brand red',
    ])
    expect(swatches.find((b) => b.classList.contains('active'))?.getAttribute('aria-label')).toBe(
      'Auto',
    )
  })

  it('a preset applies and persists; light ink darkens the scrim, dark ink lightens it', async () => {
    await render()
    await openWallpaperCenter()
    await click(swatchBy('White'))
    expect(prefSaves.at(-1)).toMatchObject({ labelColor: '#ffffff' })
    expect(labelVar('--desktop-label-ink')).toBe('#ffffff')
    expect(labelVar('--desktop-label-scrim')).toBe('var(--desktop-label-scrim-light-ink)')
    await click(swatchBy('Brand blue'))
    expect(prefSaves.at(-1)).toMatchObject({ labelColor: '#0f7fff' })
    expect(labelVar('--desktop-label-ink')).toBe('#0f7fff')
    expect(labelVar('--desktop-label-scrim')).toBe('var(--desktop-label-scrim-dark-ink)')
  })

  it('自动 resets to the theme ink and scrim', async () => {
    storedPrefs = { ...storedPrefs, labelColor: '#ffffff' }
    await render()
    expect(labelVar('--desktop-label-ink')).toBe('#ffffff')
    await openWallpaperCenter()
    await click(swatchBy('Auto'))
    expect(prefSaves.at(-1)).toMatchObject({ labelColor: '' })
    expect(labelVar('--desktop-label-ink')).toBe('')
    expect(labelVar('--desktop-label-scrim')).toBe('')
  })

  it('a dark builtin wallpaper flips the automatic ink to light', async () => {
    storedPrefs = { ...storedPrefs, wallpaper: 'builtin:starfield-night', labelColor: '' }
    await render()
    expect(labelVar('--desktop-label-ink')).toBe('#ffffff')
    expect(labelVar('--desktop-label-scrim')).toBe('var(--desktop-label-scrim-light-ink)')
  })

  it('a light builtin wallpaper keeps the theme ink under the automatic choice', async () => {
    storedPrefs = { ...storedPrefs, wallpaper: 'builtin:minimal-light', labelColor: '' }
    await render()
    expect(labelVar('--desktop-label-ink')).toBe('')
  })

  it('the HEX field applies valid values live and marks invalid ones without applying', async () => {
    await render()
    await openWallpaperCenter()
    await typeInto(hex(), '#32')
    expect(hex().classList.contains('invalid')).toBe(true)
    expect(hex().getAttribute('aria-invalid')).toBe('true')
    expect(prefSaves).toEqual([]) // marked, never applied
    await typeInto(hex(), '#3276cd')
    expect(hex().classList.contains('invalid')).toBe(false)
    expect(prefSaves.at(-1)).toMatchObject({ labelColor: '#3276cd' })
    expect(labelVar('--desktop-label-ink')).toBe('#3276cd')
    // a 3-digit hex is valid, and typing normalizes to lowercase
    await typeInto(hex(), '#F0F')
    expect(prefSaves.at(-1)).toMatchObject({ labelColor: '#f0f' })
    expect(hex().value).toBe('#f0f')
    // clearing the field resets to automatic
    await typeInto(hex(), '')
    expect(prefSaves.at(-1)).toMatchObject({ labelColor: '' })
  })

  it('presets, the native picker and the HEX field stay in sync', async () => {
    await render()
    await openWallpaperCenter()
    await click(swatchBy('Brand green'))
    expect(hex().value).toBe('#16a34a')
    const picker = host.querySelector<HTMLInputElement>('.desktop-wp-color-chip input')!
    expect(picker.value).toBe('#16a34a')
    await typeInto(picker, '#ef4444')
    expect(prefSaves.at(-1)).toMatchObject({ labelColor: '#ef4444' })
    expect(hex().value).toBe('#ef4444')
    expect(picker.value).toBe('#ef4444')
  })

  it('loads the saved label color on launch, writing nothing back', async () => {
    storedPrefs = { ...storedPrefs, labelColor: '#ffffff' }
    await render()
    expect(labelVar('--desktop-label-ink')).toBe('#ffffff')
    expect(labelVar('--desktop-label-scrim')).toBe('var(--desktop-label-scrim-light-ink)')
    expect(prefSaves).toEqual([])
  })
})
