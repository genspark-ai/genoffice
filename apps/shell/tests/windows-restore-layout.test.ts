import { beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * Windows minimize/restore lifecycle (src/main/tab-manager.ts and
 * src/main/detached-windows.ts): the shell window and detached editor windows
 * re-fit their views after a taskbar restore, reject transient degenerate
 * client bounds, and never reload, recreate, re-activate or re-focus anything
 * along the way. Electron and the app mains are faked; the handlers the real
 * constructors register are what these tests exercise.
 */

const TAB_STRIP_HEIGHT = 40

const { FakeWindow, detachedWindows } = vi.hoisted(() => {
  const WINDOW_WIDTH = 800
  const WINDOW_HEIGHT = 600

  class FakeWindow {
    handlers = new Map<string, Array<(...args: unknown[]) => void>>()
    destroyed = false
    minimized = false
    shown = true
    focused = false
    bounds = { x: 0, y: 0, width: WINDOW_WIDTH, height: WINDOW_HEIGHT }
    title = ''
    position: [number, number] = [0, 0]
    contentView = { addChildView: vi.fn(), removeChildView: vi.fn() }
    webContents = { once: vi.fn(), focus: vi.fn(), send: vi.fn() }

    on(event: string, fn: (...args: unknown[]) => void): void {
      const list = this.handlers.get(event) ?? []
      list.push(fn)
      this.handlers.set(event, list)
    }

    emit(event: string, ...args: unknown[]): void {
      for (const fn of this.handlers.get(event) ?? []) fn(...args)
    }

    isDestroyed(): boolean {
      return this.destroyed
    }
    isMinimized(): boolean {
      return this.minimized
    }
    isVisible(): boolean {
      return this.shown && !this.destroyed
    }
    isFocused(): boolean {
      return this.focused
    }
    getContentBounds(): { x: number; y: number; width: number; height: number } {
      return { ...this.bounds }
    }
    getPosition(): [number, number] {
      return this.position
    }
    setPosition(x: number, y: number): void {
      this.position = [x, y]
    }
    setTitle(title: string): void {
      this.title = title
    }
    getTitle(): string {
      return this.title
    }
    show(): void {
      this.shown = true
    }
    showInactive(): void {
      this.shown = true
    }
    focus(): void {
      this.focused = true
    }
    restore(): void {
      this.minimized = false
    }
    destroy(): void {
      this.destroyed = true
      this.emit('closed')
    }
  }

  const detachedWindows: InstanceType<typeof FakeWindow>[] = []
  return { FakeWindow, detachedWindows }
})

vi.mock('electron', () => ({
  BrowserWindow: class extends FakeWindow {
    constructor(opts: { title?: string; show?: boolean; x?: number; y?: number } = {}) {
      super()
      this.title = opts.title ?? ''
      this.shown = opts.show !== false
      this.position = [opts.x ?? 0, opts.y ?? 0]
      this.bounds = { x: 0, y: 0, width: 100, height: 100 }
      detachedWindows.push(this)
    }
  },
  screen: {
    getCursorScreenPoint: () => ({ x: 0, y: 0 }),
    getDisplayNearestPoint: () => ({ workArea: { x: 0, y: 0, width: 2000, height: 1200 } }),
  },
  shell: { openExternal: vi.fn() },
  ipcMain: { on: vi.fn(), handle: vi.fn(), removeHandler: vi.fn() },
  WebContentsView: class {
    webContents = {
      id: 900001,
      setWindowOpenHandler: () => ({ action: 'deny' as const }),
      loadURL: () => Promise.resolve(),
      focus: () => {},
      on: () => {},
      once: () => {},
      close: () => {},
      reload: () => {},
      isDestroyed: () => false,
    }
    setBounds = () => {}
    setVisible = () => {}
  },
}))

interface FakeView {
  webContents: {
    id: number
    once: ReturnType<typeof vi.fn>
    close: ReturnType<typeof vi.fn>
    reload: ReturnType<typeof vi.fn>
    focus: ReturnType<typeof vi.fn>
    isDestroyed: () => boolean
    on: ReturnType<typeof vi.fn>
    listeners: Map<string, Array<(...args: unknown[]) => void>>
  }
  setVisible: ReturnType<typeof vi.fn>
  setBounds: ReturnType<typeof vi.fn>
}

let nextId = 1
function makeView(): FakeView {
  const listeners = new Map<string, Array<(...args: unknown[]) => void>>()
  const id = nextId++
  return {
    webContents: {
      id,
      listeners,
      once: vi.fn(),
      close: vi.fn(),
      reload: vi.fn(),
      focus: vi.fn(),
      isDestroyed: () => false,
      on: vi.fn((event: string, fn: (...args: unknown[]) => void) => {
        const list = listeners.get(event) ?? []
        list.push(fn)
        listeners.set(event, list)
      }),
    },
    setVisible: vi.fn(),
    setBounds: vi.fn(),
  }
}

let lastDocsView: FakeView | null = null
let lastSheetsView: FakeView | null = null
let lastSlidesView: FakeView | null = null

const createDocsView = vi.fn(() => {
  lastDocsView = makeView()
  return lastDocsView
})
const docsQueryDirty = vi.fn(() => Promise.resolve(false))
const markDocsNewBlank = vi.fn()
const queueDocsAiContent = vi.fn()
const requestDocsClose = vi.fn(() => Promise.resolve(true))
const setActiveDocsResolver = vi.fn()
const teardownDocsRenderer = vi.fn()

vi.mock('../../docs/src/main/docs-main', () => ({
  createDocsView: (...args: unknown[]) => createDocsView(...(args as [])),
  docsQueryDirty: (...args: unknown[]) => docsQueryDirty(...(args as [])),
  markDocsNewBlank: (...args: unknown[]) => markDocsNewBlank(...args),
  queueDocsAiContent: (...args: unknown[]) => queueDocsAiContent(...args),
  requestDocsClose: (...args: unknown[]) => requestDocsClose(...(args as [])),
  setActiveDocsResolver: (...args: unknown[]) => setActiveDocsResolver(...args),
  teardownDocsRenderer: (...args: unknown[]) => teardownDocsRenderer(...args),
}))

const createSheetsView = vi.fn(() => {
  lastSheetsView = makeView()
  return lastSheetsView
})
const nudgeQueuedWorkbook = vi.fn()
const queueWorkbookForView = vi.fn()
const requestSheetsClose = vi.fn(() => Promise.resolve(true))
const setActiveSheetsWebContents = vi.fn()
const setSheetsNewBlank = vi.fn()
const sheetsPendingEditCount = vi.fn(() => 0)

vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsView: (...args: unknown[]) => createSheetsView(...(args as [])),
  nudgeQueuedWorkbook: (...args: unknown[]) => nudgeQueuedWorkbook(...args),
  queueWorkbookForView: (...args: unknown[]) => queueWorkbookForView(...args),
  requestSheetsClose: (...args: unknown[]) => requestSheetsClose(...(args as [])),
  setActiveSheetsWebContents: (...args: unknown[]) => setActiveSheetsWebContents(...args),
  setSheetsNewBlank: (...args: unknown[]) => setSheetsNewBlank(...args),
  sheetsPendingEditCount: (...args: unknown[]) => sheetsPendingEditCount(...(args as [])),
}))

const createSlidesView = vi.fn(() => {
  lastSlidesView = makeView()
  return lastSlidesView
})
const requestSlidesClose = vi.fn(() => Promise.resolve(true))
const setActiveSlidesWebContents = vi.fn()
const slidesIsDirty = vi.fn(() => false)

vi.mock('../../slides/src/main/slides-main', () => ({
  createSlidesView: (...args: unknown[]) => createSlidesView(...(args as [])),
  requestSlidesClose: (...args: unknown[]) => requestSlidesClose(...(args as [])),
  setActiveSlidesWebContents: (...args: unknown[]) => setActiveSlidesWebContents(...args),
  slidesIsDirty: (...args: unknown[]) => slidesIsDirty(...(args as [])),
}))

const createPdfView = vi.fn(() => makeView())
const clearPdfDirty = vi.fn()
const pdfIsDirty = vi.fn(() => false)
const requestPdfClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  createPdfView: (...args: unknown[]) => createPdfView(...(args as [])),
  clearPdfDirty: (...args: unknown[]) => clearPdfDirty(...args),
  pdfIsDirty: (...args: unknown[]) => pdfIsDirty(...(args as [])),
  requestPdfClose: (...args: unknown[]) => requestPdfClose(...(args as [])),
}))

const createMarkdownView = vi.fn(() => makeView())
const markdownIsDirty = vi.fn(() => false)
const requestMarkdownClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../markdown/src/main/markdown-main', () => ({
  createMarkdownView: (...args: unknown[]) => createMarkdownView(...(args as [])),
  markdownIsDirty: (...args: unknown[]) => markdownIsDirty(...(args as [])),
  requestMarkdownClose: (...args: unknown[]) => requestMarkdownClose(...(args as [])),
}))

const createHtmlView = vi.fn(() => makeView())
const createHtmlPresentView = vi.fn(() => makeView())
const htmlIsDirty = vi.fn(() => false)
const requestHtmlClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../html/src/main/html-main', () => ({
  createHtmlView: (...args: unknown[]) => createHtmlView(...(args as [])),
  createHtmlPresentView: (...args: unknown[]) => createHtmlPresentView(...(args as [])),
  htmlIsDirty: (...args: unknown[]) => htmlIsDirty(...(args as [])),
  requestHtmlClose: (...args: unknown[]) => requestHtmlClose(...(args as [])),
}))

import { TabManager } from '../src/main/tab-manager'
import { createDetachedEditorWindow } from '../src/main/detached-windows'

describe('shell window restore layout', () => {
  let shell: InstanceType<typeof FakeWindow>
  let tm: TabManager
  let onChanged: ReturnType<typeof vi.fn>
  let applyMenuFor: ReturnType<typeof vi.fn>
  let uiScale: () => number

  beforeEach(() => {
    shell = new FakeWindow()
    onChanged = vi.fn()
    applyMenuFor = vi.fn()
    uiScale = () => 1
    lastDocsView = null
    lastSheetsView = null
    lastSlidesView = null
    detachedWindows.length = 0
    tm = new TabManager(
      shell as unknown as import('electron').BrowserWindow,
      onChanged,
      applyMenuFor,
      undefined,
      () => uiScale(),
    )
  })

  // Scenario 1: Valid restore updates the active view
  it('Scenario 1: valid restore updates the active view', () => {
    tm.openDocsTab()
    const view = lastDocsView!
    view.setBounds.mockClear()

    shell.bounds = { x: 10, y: 20, width: 1024, height: 768 }
    shell.emit('restore')

    expect(view.setBounds).toHaveBeenCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: 1024,
      height: 768 - TAB_STRIP_HEIGHT,
    })
  })

  // Scenario 2: Restore with initial 0x0 bounds does not destroy previous geometry
  it('Scenario 2: restore with initial 0x0 bounds does not destroy the previous geometry', () => {
    tm.openDocsTab()
    const view = lastDocsView!
    view.setBounds.mockClear()

    shell.minimized = true
    shell.bounds = { x: 0, y: 0, width: 0, height: 0 }
    shell.emit('restore')

    expect(view.setBounds).not.toHaveBeenCalled()
  })

  // Scenario 3: Bounds settle after the restore event
  // Scenario 4: Deferred reconciliation applies settled bounds
  it('Scenario 3 & 4: bounds settle after restore event and deferred reconciliation applies settled bounds', async () => {
    tm.openDocsTab()
    const view = lastDocsView!
    view.setBounds.mockClear()

    shell.bounds = { x: 0, y: 0, width: 0, height: 0 }
    shell.emit('restore')
    expect(view.setBounds).not.toHaveBeenCalled()

    shell.bounds = { x: 0, y: 0, width: 1200, height: 800 }
    await new Promise((r) => setImmediate(r))

    expect(view.setBounds).toHaveBeenCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: 1200,
      height: 800 - TAB_STRIP_HEIGHT,
    })
  })

  // Scenario 5: Destroyed windows are safely ignored
  it('Scenario 5: destroyed windows are safely ignored', async () => {
    tm.openDocsTab()
    const view = lastDocsView!
    view.setBounds.mockClear()

    shell.destroy()
    expect(() => shell.emit('restore')).not.toThrow()
    await new Promise((r) => setImmediate(r))
    expect(view.setBounds).not.toHaveBeenCalled()
  })

  // Scenario 6: Renderer identity remains unchanged
  it('Scenario 6: renderer identity remains unchanged across restore', async () => {
    tm.openDocsTab()
    const view = lastDocsView!
    const wcBefore = view.webContents

    shell.emit('restore')
    await new Promise((r) => setImmediate(r))

    const activeTab = tm.list().find((t) => t.active)
    expect(activeTab).toBeDefined()
    expect(view.webContents).toBe(wcBefore)
  })

  // Scenario 7: Restore does not reload documents
  it('Scenario 7: restore does not reload documents', async () => {
    tm.openDocsTab()
    const view = lastDocsView!

    shell.emit('restore')
    await new Promise((r) => setImmediate(r))

    expect(view.webContents.reload).not.toHaveBeenCalled()
  })

  // Scenario 8: Restore does not switch tabs
  it('Scenario 8: restore does not switch tabs or trigger onChanged', async () => {
    tm.openDocsTab()
    const tabsBefore = tm.list()
    onChanged.mockClear()

    shell.emit('restore')
    await new Promise((r) => setImmediate(r))

    expect(tm.list()).toEqual(tabsBefore)
    expect(onChanged).not.toHaveBeenCalled()
  })

  // Scenario 9: Restore does not steal focus
  it('Scenario 9: restore does not steal focus when window is unfocused', async () => {
    tm.openDocsTab()
    const view = lastDocsView!
    shell.focused = false
    view.webContents.focus.mockClear()

    shell.emit('restore')
    await new Promise((r) => setImmediate(r))

    expect(view.webContents.focus).not.toHaveBeenCalled()
  })

  // Scenario 10: Detached editor restore works
  it('Scenario 10: detached editor window restore re-fits view and rejects degenerate bounds', async () => {
    const view = makeView()
    createDetachedEditorWindow({
      view: view as unknown as import('electron').WebContentsView,
      kind: 'docs',
      title: 'Detached Document',
      applyMenuFor: vi.fn(),
    })

    const win = detachedWindows[detachedWindows.length - 1]
    expect(win).toBeDefined()

    view.setBounds.mockClear()
    win.bounds = { x: 50, y: 50, width: 900, height: 700 }
    win.emit('restore')

    expect(view.setBounds).toHaveBeenCalledWith({
      x: 0,
      y: 0,
      width: 900,
      height: 700,
    })

    // Degenerate bounds rejected
    view.setBounds.mockClear()
    win.bounds = { x: 0, y: 0, width: 0, height: 0 }
    win.emit('restore')
    expect(view.setBounds).not.toHaveBeenCalled()
  })

  // Scenario 11: Fullscreen behavior remains correct
  it('Scenario 11: restore keeps HTML fullscreen views covering the whole window', () => {
    tm.openSlidesTab()
    const view = lastSlidesView!
    const handlers = view.webContents.listeners.get('enter-html-full-screen')
    expect(handlers).toBeDefined()
    expect(handlers!.length).toBeGreaterThan(0)
    const onEnterHtmlFullScreen = handlers![0]

    onEnterHtmlFullScreen()
    view.setBounds.mockClear()

    shell.bounds = { x: 0, y: 0, width: 1920, height: 1080 }
    shell.emit('restore')

    // Whole window coverage (no tab strip offset)
    expect(view.setBounds).toHaveBeenCalledWith({
      x: 0,
      y: 0,
      width: 1920,
      height: 1080,
    })
  })

  // Scenario 12: UI scaling remains correct
  it('Scenario 12: UI scaling remains correct on restore', () => {
    uiScale = () => 1.5
    tm.openDocsTab()
    const view = lastDocsView!
    view.setBounds.mockClear()

    shell.bounds = { x: 0, y: 0, width: 1000, height: 800 }
    shell.emit('restore')

    const scaledStrip = Math.round(TAB_STRIP_HEIGHT * 1.5) // 60
    expect(view.setBounds).toHaveBeenCalledWith({
      x: 0,
      y: scaledStrip,
      width: 1000,
      height: 800 - scaledStrip,
    })
  })

  // Scenario 13: Existing Linux/X11 resize behavior still works
  it('Scenario 13: existing Linux/X11 resize behavior still works', async () => {
    tm.openDocsTab()
    const view = lastDocsView!
    view.setBounds.mockClear()

    shell.bounds = { x: 0, y: 0, width: 900, height: 600 }
    shell.emit('resize')

    expect(view.setBounds).toHaveBeenCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: 900,
      height: 600 - TAB_STRIP_HEIGHT,
    })

    view.setBounds.mockClear()
    shell.bounds = { x: 0, y: 0, width: 1100, height: 700 }
    await new Promise((r) => setImmediate(r))

    expect(view.setBounds).toHaveBeenCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: 1100,
      height: 700 - TAB_STRIP_HEIGHT,
    })
  })

  // Scenario 14: Existing tab activation tests still pass
  it('Scenario 14: activating a tab while window reports degenerate bounds does not zero-size it', () => {
    const docId = tm.openDocsTab()
    const view = lastDocsView!

    shell.bounds = { x: 0, y: 0, width: 0, height: 0 }
    view.setBounds.mockClear()

    tm.activateTab(docId)
    expect(view.setBounds).not.toHaveBeenCalled()
  })
})
