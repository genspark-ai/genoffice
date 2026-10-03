import { beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('node:fs', () => {
  const realpathSync = ((path?: string) => {
    if (path === undefined) return undefined
    if (path === '/real/file' || path === '/canonical/file' || path === '/REAL/FILE')
      return '/real/file'
    return path
  }) as typeof import('node:fs').realpathSync

  realpathSync.native = ((path?: string) => {
    if (path === undefined) return undefined
    if (path === '/real/file' || path === '/canonical/file' || path === '/REAL/FILE')
      return '/real/file'
    return path
  }) as typeof import('node:fs').realpathSync.native

  return { realpathSync }
})

/**
 * TabManager (src/main/tab-manager.ts): tab list state, activation,
 * close guards, and view lifecycle inside the shell's single window.
 * Electron and the per-module main entrypoints are mocked; only the
 * manager's own observable behavior is asserted.
 */

interface FakeWebContents {
  id: number
  on: ReturnType<typeof vi.fn>
  once: ReturnType<typeof vi.fn>
  close: ReturnType<typeof vi.fn>
  reload: ReturnType<typeof vi.fn>
  focus: ReturnType<typeof vi.fn>
  isDestroyed: ReturnType<typeof vi.fn>
  listeners: Map<string, () => void>
}

interface FakeView {
  webContents: FakeWebContents
  setVisible: ReturnType<typeof vi.fn>
  setBounds: ReturnType<typeof vi.fn>
}

let nextWebContentsId = 1

function makeFakeView(): FakeView {
  const listeners = new Map<string, () => void>()
  return {
    webContents: {
      id: nextWebContentsId++,
      listeners,
      on: vi.fn((event: string, handler: () => void) => {
        listeners.set(event, handler)
      }),
      once: vi.fn(),
      close: vi.fn(),
      reload: vi.fn(),
      focus: vi.fn(),
      isDestroyed: vi.fn(() => false),
    },
    setVisible: vi.fn(),
    setBounds: vi.fn(),
  }
}

vi.mock('electron', () => ({ BrowserWindow: class {} }))

const createDocsView = vi.fn(() => makeFakeView())
const docsQueryDirty = vi.fn(() => Promise.resolve(false))
const markDocsNewBlank = vi.fn()
const requestDocsClose = vi.fn(() => Promise.resolve(true))
const setActiveDocsResolver = vi.fn()
const teardownDocsRenderer = vi.fn()

vi.mock('../../docs/src/main/docs-main', () => ({
  createDocsView: (...args: unknown[]) => createDocsView(...(args as [])),
  docsQueryDirty: (...args: unknown[]) => docsQueryDirty(...(args as [])),
  markDocsNewBlank: (...args: unknown[]) => markDocsNewBlank(...args),
  requestDocsClose: (...args: unknown[]) => requestDocsClose(...(args as [])),
  setActiveDocsResolver: (...args: unknown[]) => setActiveDocsResolver(...args),
  teardownDocsRenderer: (...args: unknown[]) => teardownDocsRenderer(...args),
}))

const createPdfView = vi.fn(() => makeFakeView())
const pdfIsDirty = vi.fn(() => false)
const clearPdfDirty = vi.fn()
const requestPdfClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  createPdfView: (...args: unknown[]) => createPdfView(...(args as [])),
  pdfIsDirty: (...args: unknown[]) => pdfIsDirty(...(args as [])),
  clearPdfDirty: (...args: unknown[]) => clearPdfDirty(...(args as [])),
  requestPdfClose: (...args: unknown[]) => requestPdfClose(...(args as [])),
}))

const createSheetsView = vi.fn(() => makeFakeView())
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

const createSlidesView = vi.fn(() => makeFakeView())
const requestSlidesClose = vi.fn(() => Promise.resolve(true))
const setActiveSlidesWebContents = vi.fn()
const slidesIsDirty = vi.fn(() => false)

vi.mock('../../slides/src/main/slides-main', () => ({
  createSlidesView: (...args: unknown[]) => createSlidesView(...(args as [])),
  requestSlidesClose: (...args: unknown[]) => requestSlidesClose(...(args as [])),
  setActiveSlidesWebContents: (...args: unknown[]) => setActiveSlidesWebContents(...args),
  slidesIsDirty: (...args: unknown[]) => slidesIsDirty(...(args as [])),
}))

const createMarkdownView = vi.fn(() => makeFakeView())
const markdownIsDirty = vi.fn(() => false)
const requestMarkdownClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../markdown/src/main/markdown-main', () => ({
  createMarkdownView: (...args: unknown[]) => createMarkdownView(...(args as [])),
  markdownIsDirty: (...args: unknown[]) => markdownIsDirty(...(args as [])),
  requestMarkdownClose: (...args: unknown[]) => requestMarkdownClose(...(args as [])),
}))

const createHtmlView = vi.fn(() => makeFakeView())
const createHtmlPresentView = vi.fn(() => makeFakeView())
const htmlIsDirty = vi.fn(() => false)
const requestHtmlClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../html/src/main/html-main', () => ({
  createHtmlView: (...args: unknown[]) => createHtmlView(...(args as [])),
  createHtmlPresentView: (...args: unknown[]) => createHtmlPresentView(...(args as [])),
  htmlIsDirty: (...args: unknown[]) => htmlIsDirty(...(args as [])),
  requestHtmlClose: (...args: unknown[]) => requestHtmlClose(...(args as [])),
}))

import { TabManager } from '../src/main/tab-manager'
import type { TabKind } from '../src/shared/tabs-api'

const TAB_STRIP_HEIGHT = 40
const WINDOW_WIDTH = 800
const WINDOW_HEIGHT = 600

interface FakeShellWindow {
  on: ReturnType<typeof vi.fn>
  webContents: { once: ReturnType<typeof vi.fn>; focus: ReturnType<typeof vi.fn> }
  isDestroyed: ReturnType<typeof vi.fn>
  isFocused: ReturnType<typeof vi.fn>
  getContentBounds: () => { x: number; y: number; width: number; height: number }
  contentView: {
    addChildView: ReturnType<typeof vi.fn>
    removeChildView: ReturnType<typeof vi.fn>
  }
}

function makeShellWindow(): FakeShellWindow {
  return {
    on: vi.fn(),
    webContents: { once: vi.fn(), focus: vi.fn() },
    isDestroyed: vi.fn(() => false),
    isFocused: vi.fn(() => true),
    getContentBounds: () => ({ x: 0, y: 0, width: WINDOW_WIDTH, height: WINDOW_HEIGHT }),
    contentView: { addChildView: vi.fn(), removeChildView: vi.fn() },
  }
}

let shellWindow: FakeShellWindow
let onChanged: ReturnType<typeof vi.fn>
let applyMenuFor: ReturnType<typeof vi.fn>
let manager: TabManager

function lastCreatedView(factory: ReturnType<typeof vi.fn>): FakeView {
  return factory.mock.results.at(-1)!.value as FakeView
}

beforeEach(() => {
  vi.clearAllMocks()
  nextWebContentsId = 1
  docsQueryDirty.mockImplementation(() => Promise.resolve(false))
  requestDocsClose.mockImplementation(() => Promise.resolve(true))
  pdfIsDirty.mockImplementation(() => false)
  sheetsPendingEditCount.mockImplementation(() => 0)
  slidesIsDirty.mockImplementation(() => false)
  shellWindow = makeShellWindow()
  onChanged = vi.fn()
  applyMenuFor = vi.fn()
  manager = new TabManager(
    shellWindow as never,
    () => onChanged(),
    (kind) => applyMenuFor(kind),
  )
})

describe('initial state', () => {
  it('starts with only the non-closable, active Home tab', () => {
    expect(manager.list()).toEqual([
      { id: 'home', kind: 'home', title: 'GenOffice', closable: false, active: true },
    ])
  })
})

describe('opening tabs', () => {
  it('opens a docs tab, activates it, and attaches its view to the window', () => {
    const id = manager.openDocsTab()
    const tabs = manager.list()
    expect(tabs).toHaveLength(2)
    expect(tabs[1]).toMatchObject({
      id,
      kind: 'docs',
      title: 'GenOffice Docs',
      closable: true,
      active: true,
    })
    expect(tabs[0].active).toBe(false)
    expect(shellWindow.contentView.addChildView).toHaveBeenCalledTimes(1)
    expect(applyMenuFor).toHaveBeenLastCalledWith('docs')
    expect(onChanged).toHaveBeenCalled()
  })

  it('titles file-backed tabs with the file basename', () => {
    manager.openDocsTab('/tmp/report.docx')
    manager.openSheetsTab('/tmp/budget.xlsx')
    manager.openSlidesTab('/tmp/deck.pptx')
    manager.openPdfTab('/tmp/scan.pdf')
    expect(manager.list().map((t) => t.title)).toEqual([
      'GenOffice',
      'report.docx',
      'budget.xlsx',
      'deck.pptx',
      'scan.pdf',
    ])
  })

  it('uses module default titles for pathless tabs', () => {
    manager.openSheetsTab()
    manager.openSlidesTab()
    expect(manager.list().map((t) => t.title)).toEqual(['GenOffice', 'AI Sheets', 'AI Slides'])
  })

  it('assigns unique, monotonic tab ids', () => {
    const a = manager.openDocsTab()
    const b = manager.openSheetsTab()
    expect(a).not.toBe(b)
    expect(a).toBe('t1')
    expect(b).toBe('t2')
  })

  it('forwards the new-blank flag to the module', () => {
    manager.openDocsTab(undefined, { newBlank: true })
    expect(markDocsNewBlank).toHaveBeenCalledTimes(1)
    manager.openSheetsTab(undefined, { newBlank: true })
    expect(setSheetsNewBlank).toHaveBeenCalledTimes(1)
  })
})

describe('spare sheets view', () => {
  it('does not warm Sheets during Home or Docs sessions', () => {
    vi.useFakeTimers()
    try {
      const homeLoaded = shellWindow.webContents.once.mock.calls.find(
        ([event]) => event === 'did-finish-load',
      )
      ;(homeLoaded?.[1] as (() => void) | undefined)?.()
      vi.advanceTimersByTime(5000)
      manager.openDocsTab('/tmp/report.docx')
      vi.advanceTimersByTime(5000)
      expect(createSheetsView).not.toHaveBeenCalled()
    } finally {
      vi.useRealTimers()
    }
  })

  it('warms a spare while Sheets is active and hands it to the next open', () => {
    vi.useFakeTimers()
    try {
      manager.openSheetsTab('/tmp/first.xlsx')
      expect(createSheetsView).toHaveBeenCalledTimes(1)
      vi.advanceTimersByTime(3000)
      expect(createSheetsView).toHaveBeenCalledTimes(2)
      const spare = lastCreatedView(createSheetsView)
      expect(shellWindow.contentView.addChildView).toHaveBeenCalledWith(spare)
      expect(spare.setVisible).toHaveBeenLastCalledWith(false)
      expect(manager.list()).toHaveLength(2)

      manager.openSheetsTab('/tmp/budget.xlsx')
      expect(createSheetsView).toHaveBeenCalledTimes(2)
      expect(queueWorkbookForView).toHaveBeenCalledWith(spare.webContents, '/tmp/budget.xlsx')
      expect(nudgeQueuedWorkbook).toHaveBeenCalledWith(spare.webContents)
      expect(spare.setVisible).toHaveBeenLastCalledWith(true)
      expect(manager.list()[2]).toMatchObject({
        kind: 'sheets',
        title: 'budget.xlsx',
        active: true,
      })

      vi.advanceTimersByTime(3000)
      expect(createSheetsView).toHaveBeenCalledTimes(3)
      expect(lastCreatedView(createSheetsView)).not.toBe(spare)
      expect(setActiveSheetsWebContents).toHaveBeenLastCalledWith(spare.webContents)
    } finally {
      vi.useRealTimers()
    }
  })

  it('creates a fresh view when no spare is ready and does not nudge it', () => {
    manager.openSheetsTab('/tmp/budget.xlsx')
    expect(createSheetsView).toHaveBeenCalledTimes(1)
    expect(createSheetsView).toHaveBeenCalledWith({
      includeAiHandlers: false,
      openingWorkbook: true,
    })
    expect(nudgeQueuedWorkbook).not.toHaveBeenCalled()
  })

  it('starts a new blank view without the opening state', () => {
    manager.openSheetsTab()
    expect(createSheetsView).toHaveBeenCalledWith({
      includeAiHandlers: false,
      openingWorkbook: false,
    })
  })

  it('drops a spare whose renderer died instead of handing it out', () => {
    vi.useFakeTimers()
    try {
      manager.openSheetsTab()
      vi.advanceTimersByTime(3000)
      const spare = lastCreatedView(createSheetsView)
      const gone = spare.webContents.once.mock.calls.find(
        ([event]) => event === 'render-process-gone',
      )
      ;(gone![1] as () => void)()
      expect(spare.webContents.close).toHaveBeenCalledTimes(1)
      manager.openSheetsTab()
      expect(createSheetsView).toHaveBeenCalledTimes(3)
    } finally {
      vi.useRealTimers()
    }
  })

  it('cancels a pending spare and closes an existing spare on leaving Sheets', () => {
    vi.useFakeTimers()
    try {
      const sheetsId = manager.openSheetsTab()
      manager.openDocsTab()
      vi.advanceTimersByTime(3000)
      expect(createSheetsView).toHaveBeenCalledTimes(1)

      manager.activateTab(sheetsId)
      vi.advanceTimersByTime(3000)
      expect(createSheetsView).toHaveBeenCalledTimes(2)
      const spare = lastCreatedView(createSheetsView)
      manager.activateTab('home')
      expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(spare)
      expect(spare.webContents.close).toHaveBeenCalledTimes(1)
    } finally {
      vi.useRealTimers()
    }
  })

  it('stays off under GENOFFICE_NO_SPARE_VIEW', () => {
    vi.stubEnv('GENOFFICE_NO_SPARE_VIEW', '1')
    vi.useFakeTimers()
    try {
      manager.openSheetsTab()
      vi.advanceTimersByTime(5000)
      expect(createSheetsView).toHaveBeenCalledTimes(1)
    } finally {
      vi.useRealTimers()
      vi.unstubAllEnvs()
    }
  })
})

describe('activation', () => {
  it('shows only the activated tab view and lays it out below the tab strip', () => {
    const docsId = manager.openDocsTab()
    const docsView = lastCreatedView(createDocsView)
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)

    manager.activateTab(docsId)
    expect(docsView.setVisible).toHaveBeenLastCalledWith(true)
    expect(sheetsView.setVisible).toHaveBeenLastCalledWith(false)
    expect(docsView.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })
    expect(manager.list().find((t) => t.id === docsId)?.active).toBe(true)
  })

  it('hands keyboard focus to the activated view so typing works right after open/switch', () => {
    const docsId = manager.openDocsTab()
    const docsView = lastCreatedView(createDocsView)
    expect(docsView.webContents.focus).toHaveBeenCalled()

    docsView.webContents.focus.mockClear()
    manager.activateTab('home')
    expect(shellWindow.webContents.focus).toHaveBeenCalled()
    manager.activateTab(docsId)
    expect(docsView.webContents.focus).toHaveBeenCalledTimes(1)
  })

  it('does not steal OS focus for a background open (window unfocused)', () => {
    shellWindow.isFocused.mockReturnValue(false)
    manager.openDocsTab()
    const docsView = lastCreatedView(createDocsView)
    expect(docsView.webContents.focus).not.toHaveBeenCalled()

    // the window `focus` handler runs it once the user comes back
    shellWindow.isFocused.mockReturnValue(true)
    manager.focusActiveView()
    expect(docsView.webContents.focus).toHaveBeenCalledTimes(1)
  })

  it('ignores activation of unknown tab ids', () => {
    onChanged.mockClear()
    manager.activateTab('nope')
    expect(onChanged).not.toHaveBeenCalled()
    expect(manager.list()[0].active).toBe(true)
  })

  it('routes the active webContents to the matching module', () => {
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)
    expect(setActiveSheetsWebContents).toHaveBeenLastCalledWith(sheetsView.webContents)

    manager.openSlidesTab()
    const slidesView = lastCreatedView(createSlidesView)
    expect(setActiveSlidesWebContents).toHaveBeenLastCalledWith(slidesView.webContents)
  })

  it('lets the active view cover the tab strip during HTML fullscreen', () => {
    manager.openSlidesTab()
    const view = lastCreatedView(createSlidesView)
    view.webContents.listeners.get('enter-html-full-screen')!()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: 0,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT,
    })
    view.webContents.listeners.get('leave-html-full-screen')!()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })
  })
})

describe('window resize layout', () => {
  function resizeHandler(): () => void {
    const call = shellWindow.on.mock.calls.find((c) => c[0] === 'resize')
    expect(call).toBeDefined()
    return call![1] as () => void
  }

  it('re-lays out after resize bounds settle (Linux/X11 stale getContentBounds)', async () => {
    // On X11, `resize` fires before the WM applies maximize bounds, so the first
    // layout still sees the pre-maximize size. The deferred layout must pick up
    // the real size on the next turn (see issue #15).
    manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    view.setBounds.mockClear()

    let width = WINDOW_WIDTH
    let height = WINDOW_HEIGHT
    shellWindow.getContentBounds = () => ({ x: 0, y: 0, width, height })

    resizeHandler()()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })

    // Bounds update after the synchronous layout, as on X11 maximize.
    width = 1920
    height = 1080
    await new Promise<void>((resolve) => setImmediate(resolve))

    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: 1920,
      height: 1080 - TAB_STRIP_HEIGHT,
    })
    expect(view.setBounds).toHaveBeenCalledTimes(2)
  })

  it('skips deferred layout after the shell window is destroyed', async () => {
    manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    view.setBounds.mockClear()

    resizeHandler()()
    expect(view.setBounds).toHaveBeenCalledTimes(1)

    shellWindow.isDestroyed.mockReturnValue(true)
    await new Promise<void>((resolve) => setImmediate(resolve))
    expect(view.setBounds).toHaveBeenCalledTimes(1)
  })
})

describe('closing tabs', () => {
  it('never closes the Home tab', async () => {
    await manager.closeTab('home')
    expect(manager.list()).toHaveLength(1)

    manager.openHomeTab()
    manager.closeActiveTab()
    await Promise.resolve()
    expect(manager.list()).toHaveLength(1)
  })

  it('removes a clean tab and falls back to the previous tab', async () => {
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)
    const slidesId = manager.openSlidesTab()

    await manager.closeTab(slidesId)
    const tabs = manager.list()
    expect(tabs.map((t) => t.id)).toEqual(['home', 't1'])
    expect(tabs[1].active).toBe(true)
    expect(sheetsView.setVisible).toHaveBeenLastCalledWith(true)
  })

  it('keeps the current tab active when closing a background tab', async () => {
    const sheetsId = manager.openSheetsTab()
    const slidesId = manager.openSlidesTab()
    await manager.closeTab(sheetsId)
    expect(manager.list().find((t) => t.id === slidesId)?.active).toBe(true)
  })

  it('detaches and destroys non-docs views on close', async () => {
    const id = manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    await manager.closeTab(id)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    expect(view.webContents.close).toHaveBeenCalledTimes(1)
  })

  it('detaches docs views without destroying the webContents (freeze workaround)', async () => {
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    await manager.closeTab(id)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    expect(view.webContents.close).not.toHaveBeenCalled()
    // the orphaned renderer must be told to go inert (recovery-copy resurrection guard)
    expect(teardownDocsRenderer).toHaveBeenCalledWith(view.webContents)
  })

  it('closes a clean docs tab after the async dirty query says clean', async () => {
    const id = manager.openDocsTab()
    await manager.closeTab(id)
    expect(docsQueryDirty).toHaveBeenCalledTimes(1)
    expect(requestDocsClose).not.toHaveBeenCalled()
    expect(manager.list()).toHaveLength(1)
  })

  it('keeps a dirty docs tab open when the user cancels the close guard', async () => {
    docsQueryDirty.mockImplementation(() => Promise.resolve(true))
    requestDocsClose.mockImplementation(() => Promise.resolve(false))
    const id = manager.openDocsTab()
    await manager.closeTab(id)
    expect(requestDocsClose).toHaveBeenCalledTimes(1)
    expect(manager.list().map((t) => t.id)).toEqual(['home', id])
  })

  it('activates a dirty background tab before showing its close guard', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    requestSheetsClose.mockImplementation(() => Promise.resolve(false))
    const sheetsId = manager.openSheetsTab()
    manager.openSlidesTab()

    await manager.closeTab(sheetsId)
    expect(requestSheetsClose).toHaveBeenCalledTimes(1)
    // the guarded tab was brought into view for the prompt
    expect(manager.list().find((t) => t.id === sheetsId)?.active).toBe(true)
  })

  it('closes a dirty sheets tab when the guard resolves true', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    requestSheetsClose.mockImplementation(() => Promise.resolve(true))
    const id = manager.openSheetsTab()
    await manager.closeTab(id)
    expect(manager.list()).toHaveLength(1)
  })

  it('does not stack close guards while one prompt is pending', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    let resolveGuard!: (ok: boolean) => void
    requestSheetsClose.mockImplementation(
      () =>
        new Promise<boolean>((resolve) => {
          resolveGuard = resolve
        }),
    )
    const id = manager.openSheetsTab()
    const first = manager.closeTab(id)
    const second = manager.closeTab(id)
    resolveGuard(true)
    await Promise.all([first, second])
    expect(requestSheetsClose).toHaveBeenCalledTimes(1)
    expect(manager.list()).toHaveLength(1)
  })
})

describe('file path bookkeeping', () => {
  it('updates the tab title when a module opens a file in an existing tab', () => {
    manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    manager.setTabFileFor(view.webContents.id, '/tmp/final.docx')
    expect(manager.list()[1].title).toBe('final.docx')
    expect(manager.findDocsTabByPath('/tmp/final.docx')).toBe('t1')
  })

  it('ignores setTabFileFor for unknown webContents', () => {
    onChanged.mockClear()
    manager.setTabFileFor(999, '/tmp/x.docx')
    expect(onChanged).not.toHaveBeenCalled()
  })

  it('renames matching tabs and reports the affected views', () => {
    manager.openDocsTab('/tmp/old.docx')
    const view = lastCreatedView(createDocsView)
    const affected = manager.renameTabFile('/tmp/old.docx', '/tmp/new.docx')
    expect(affected).toEqual([{ kind: 'docs', webContents: view.webContents }])
    expect(manager.list()[1].title).toBe('new.docx')
    expect(manager.findDocsTabByPath('/tmp/new.docx')).toBe('t1')
    expect(manager.findDocsTabByPath('/tmp/old.docx')).toBeUndefined()
  })

  it('returns no affected views when nothing matches a rename', () => {
    onChanged.mockClear()
    expect(manager.renameTabFile('/tmp/none.docx', '/tmp/new.docx')).toEqual([])
    expect(onChanged).not.toHaveBeenCalled()
  })

  /**
   * One case per TabKind a file can live in. The shell dispatches a rename to a
   * per-app hook on the kind renameTabFile reports, so a kind with no case here
   * is a kind whose editor would silently keep saving to the old path. `home`
   * carries no file, so it gets the negative case below instead of a positive
   * one.
   */
  const RENAMABLE: {
    kind: Exclude<TabKind, 'home'>
    file: string
    renamed: string
    open: () => FakeView
  }[] = [
    {
      kind: 'docs',
      file: '/tmp/one.docx',
      renamed: '/tmp/two.docx',
      open: () => {
        manager.openDocsTab('/tmp/one.docx')
        return lastCreatedView(createDocsView)
      },
    },
    {
      kind: 'sheets',
      file: '/tmp/one.xlsx',
      renamed: '/tmp/two.xlsx',
      open: () => {
        manager.openSheetsTab('/tmp/one.xlsx')
        return lastCreatedView(createSheetsView)
      },
    },
    {
      kind: 'slides',
      file: '/tmp/one.pptx',
      renamed: '/tmp/two.pptx',
      open: () => {
        manager.openSlidesTab('/tmp/one.pptx')
        return lastCreatedView(createSlidesView)
      },
    },
    {
      kind: 'pdf',
      file: '/tmp/one.pdf',
      renamed: '/tmp/two.pdf',
      open: () => {
        manager.openPdfTab('/tmp/one.pdf')
        return lastCreatedView(createPdfView)
      },
    },
    {
      kind: 'markdown',
      file: '/tmp/one.md',
      renamed: '/tmp/two.md',
      open: () => {
        manager.openMarkdownTab('/tmp/one.md')
        return lastCreatedView(createMarkdownView)
      },
    },
    {
      kind: 'html',
      file: '/tmp/one.html',
      renamed: '/tmp/two.html',
      open: () => {
        manager.openHtmlTab('/tmp/one.html')
        return lastCreatedView(createHtmlView)
      },
    },
  ]

  it.each(RENAMABLE)('a rename reports the $kind view', (c) => {
    const view = c.open()
    expect(manager.renameTabFile(c.file, c.renamed)).toEqual([
      { kind: c.kind, webContents: view.webContents },
    ])
  })

  it.each(RENAMABLE)('a rename retitles and repaths the $kind tab', (c) => {
    c.open()
    manager.renameTabFile(c.file, c.renamed)
    const tab = manager.list().find((t) => t.id === 't1')
    expect(tab?.filePath).toBe(c.renamed)
    expect(tab?.title).toBe('two' + c.file.slice(c.file.lastIndexOf('.')))
  })

  it('never claims the home tab, which has no file on disk', () => {
    const before = manager.list().map((t) => ({ id: t.id, title: t.title, filePath: t.filePath }))
    expect(before.some((t) => t.id === 'home')).toBe(true)
    expect(manager.renameTabFile('/tmp/one.docx', '/tmp/two.docx')).toEqual([])
    expect(manager.list().map((t) => ({ id: t.id, title: t.title, filePath: t.filePath }))).toEqual(
      before,
    )
  })

  it('covers every kind the manager can open', () => {
    // The table is checked against what the manager can actually produce rather
    // than a hand-written list, so a new open*Tab method cannot ship without a
    // rename case. A Record<TabKind, true> would read better, but apps/shell
    // tsconfig includes src/** only, so nothing under tests/ ever reaches the
    // typecheck CI runs.
    const openers = Object.getOwnPropertyNames(TabManager.prototype).filter((name) =>
      /^open[A-Z]\w*Tab$/.test(name),
    )
    expect(openers.length).toBeGreaterThanOrEqual(7)
    const opened: string[] = []
    for (const name of openers) {
      const before = manager.list().length
      const opener = (manager as unknown as Record<string, (arg?: unknown) => unknown>)[name]
      opener?.call(manager, '/tmp/probe.probe')
      const after = manager.list()
      if (after.length > before) opened.push(after[after.length - 1]!.kind)
    }
    expect([...new Set(opened)].sort()).toEqual(RENAMABLE.map((c) => c.kind).sort())
  })

  it('finds tabs by kind and path', () => {
    manager.openSheetsTab('/tmp/a.xlsx')
    manager.openSlidesTab('/tmp/b.pptx')
    manager.openPdfTab('/tmp/c.pdf')
    expect(manager.findSheetsTab()).toBe('t1')
    expect(manager.findSheetsTabByPath('/tmp/a.xlsx')).toBe('t1')
    expect(manager.findSlidesTabByPath('/tmp/b.pptx')).toBe('t2')
    expect(manager.findPdfTabByPath('/tmp/c.pdf')).toBe('t3')
    expect(manager.findPdfTabByPath('/tmp/missing.pdf')).toBeUndefined()
  })

  it('finds every document family by a canonicalized path alias', () => {
    const docsId = manager.openDocsTab('/real/file')
    const sheetsId = manager.openSheetsTab('/real/file')
    const slidesId = manager.openSlidesTab('/real/file')
    const pdfId = manager.openPdfTab('/real/file')
    // markdown/html view factories need electron protocol mocks the shell
    // suite does not provide, so seed their tab records directly: the
    // finders only read kind/view/filePath.
    const seedTab = (kind: string, filePath: string) => {
      const tabs = (
        manager as unknown as {
          tabs: Array<{ id: string; kind: string; view: unknown; filePath: string }>
        }
      ).tabs
      const id = `seed-${kind}`
      tabs.push({ id, kind, view: {}, filePath })
      return id
    }
    const markdownId = seedTab('markdown', '/real/file')
    const htmlId = seedTab('html', '/real/file')

    expect(manager.findDocsTabByPath('/canonical/file')).toBe(docsId)
    expect(manager.findSheetsTabByPath('/REAL/FILE')).toBe(sheetsId)
    expect(manager.findSlidesTabByPath('/canonical/file')).toBe(slidesId)
    expect(manager.findPdfTabByPath('/REAL/FILE')).toBe(pdfId)
    expect(manager.findMarkdownTabByPath('/canonical/file')).toBe(markdownId)
    expect(manager.findHtmlTabByPath('/REAL/FILE')).toBe(htmlId)
    expect(manager.findDocsTabByPath('/missing')).toBeUndefined()
    expect(manager.findDocsTabByPath()).toBeUndefined()
  })

  it('reloads an existing pdf tab so a re-export rereads the file from disk', () => {
    const id = manager.openPdfTab('/tmp/c.pdf')
    const view = lastCreatedView(createPdfView)
    manager.reloadTab(id)
    expect(clearPdfDirty).toHaveBeenCalledWith(view.webContents.id)
    expect(view.webContents.reload).toHaveBeenCalledTimes(1)
  })

  it('reports the active pdf tab with its id (so callers can re-activate it)', () => {
    const pdfId = manager.openPdfTab('/tmp/c.pdf')
    const active = manager.activePdfTab()
    expect(active?.id).toBe(pdfId)
    expect(active?.filePath).toBe('/tmp/c.pdf')
    manager.openDocsTab()
    expect(manager.activePdfTab()).toBeUndefined()
  })
})

describe('dirty-tab queries (shell close guard)', () => {
  it('lists only tabs whose module reports unsaved changes', () => {
    const dirtySheetsId = manager.openSheetsTab()
    const dirtyView = lastCreatedView(createSheetsView)
    manager.openSheetsTab()
    sheetsPendingEditCount.mockImplementation((id: number) =>
      id === dirtyView.webContents.id ? 3 : 0,
    )

    const dirty = manager.dirtySheetsTabs()
    expect(dirty).toEqual([{ id: dirtySheetsId, webContents: dirtyView.webContents }])
  })

  it('lists dirty pdf and slides tabs', () => {
    const pdfId = manager.openPdfTab('/tmp/c.pdf')
    const slidesId = manager.openSlidesTab()
    expect(manager.dirtyPdfTabs()).toEqual([])
    expect(manager.dirtySlidesTabs()).toEqual([])
    pdfIsDirty.mockImplementation(() => true)
    slidesIsDirty.mockImplementation(() => true)
    expect(manager.dirtyPdfTabs().map((t) => t.id)).toEqual([pdfId])
    expect(manager.dirtySlidesTabs().map((t) => t.id)).toEqual([slidesId])
  })

  it('lists every live docs tab for the async dirtiness sweep', () => {
    manager.openDocsTab()
    manager.openSheetsTab()
    manager.openDocsTab('/tmp/a.docx')
    expect(manager.docsTabs().map((t) => t.id)).toEqual(['t1', 't3'])
  })
})

describe('detach / attach (Open in New Window, tear-off, dock)', () => {
  it('lifts a tab out without closing its renderer and hands the live view back', () => {
    manager.openDocsTab('/tmp/a.docx')
    const id = manager.openPdfTab('/tmp/scan.pdf')
    const view = lastCreatedView(createPdfView)
    const record = manager.detachTab(id)
    expect(record).toMatchObject({ kind: 'pdf', title: 'scan.pdf', filePath: '/tmp/scan.pdf' })
    expect(record!.view).toBe(view)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    expect(view.webContents.close).not.toHaveBeenCalled()
    expect(manager.list().map((t) => t.id)).toEqual(['home', 't1'])
    expect(manager.list()[1].active).toBe(true)
  })

  it('every document kind can be detached; Home and Present tabs cannot', () => {
    const docs = manager.openDocsTab()
    const sheets = manager.openSheetsTab()
    const slides = manager.openSlidesTab()
    const pdf = manager.openPdfTab('/tmp/scan.pdf')
    const markdown = manager.openMarkdownTab()
    const html = manager.openHtmlTab()
    const present = manager.openHtmlPresentTab({ id: 999 } as never, 'Preview')
    for (const id of [docs, sheets, slides, pdf, markdown, html])
      expect(manager.canDetachTab(id)).toBe(true)
    expect(manager.canDetachTab('home')).toBe(false)
    expect(manager.canDetachTab(present)).toBe(false)
    expect(manager.canDetachTab('nope')).toBe(false)
    expect(manager.detachTab(present)).toBeNull()
  })

  it('attaches a view back as a new, active tab at the requested slot', () => {
    manager.openDocsTab('/tmp/a.docx')
    manager.openDocsTab('/tmp/b.docx')
    const id = manager.openSlidesTab('/tmp/deck.pptx')
    const record = manager.detachTab(id)!
    shellWindow.contentView.addChildView.mockClear()

    const newId = manager.attachTab(record, 1)
    expect(newId).not.toBe(id)
    expect(shellWindow.contentView.addChildView).toHaveBeenCalledWith(record.view)
    expect(record.view.setVisible).toHaveBeenLastCalledWith(true)
    expect(record.view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })
    expect(manager.list().map((t) => [t.id, t.title, t.active])).toEqual([
      ['home', 'GenOffice', false],
      [newId, 'deck.pptx', true],
      ['t1', 'a.docx', false],
      ['t2', 'b.docx', false],
    ])
    expect(setActiveSlidesWebContents).toHaveBeenLastCalledWith(record.view.webContents)
    expect(applyMenuFor).toHaveBeenLastCalledWith('slides')
  })

  it('keeps Home pinned: slot 0 and negative slots land right after it, out-of-range appends', () => {
    manager.openDocsTab('/tmp/a.docx')
    const first = manager.detachTab(manager.openSheetsTab('/tmp/x.xlsx'))!
    const second = manager.detachTab(manager.openSheetsTab('/tmp/y.xlsx'))!
    const third = manager.detachTab(manager.openSheetsTab('/tmp/z.xlsx'))!
    manager.attachTab(first, 0)
    expect(manager.list().map((t) => t.title)).toEqual(['GenOffice', 'x.xlsx', 'a.docx'])
    manager.attachTab(second, -4)
    expect(manager.list().map((t) => t.title)).toEqual(['GenOffice', 'y.xlsx', 'x.xlsx', 'a.docx'])
    manager.attachTab(third, 99)
    expect(manager.list().map((t) => t.title)).toEqual([
      'GenOffice',
      'y.xlsx',
      'x.xlsx',
      'a.docx',
      'z.xlsx',
    ])
  })

  it('appends when no slot is given', () => {
    manager.openDocsTab('/tmp/a.docx')
    const record = manager.detachTab(manager.openPdfTab('/tmp/scan.pdf'))!
    manager.openDocsTab('/tmp/b.docx')
    manager.attachTab(record)
    expect(manager.list().map((t) => t.title)).toEqual([
      'GenOffice',
      'a.docx',
      'b.docx',
      'scan.pdf',
    ])
  })

  it('installs the HTML-fullscreen listeners once per view across detach and re-attach', () => {
    const id = manager.openSlidesTab('/tmp/deck.pptx')
    const view = lastCreatedView(createSlidesView)
    const fullScreenRegistrations = () =>
      view.webContents.on.mock.calls.filter(([event]) => event === 'enter-html-full-screen').length
    expect(fullScreenRegistrations()).toBe(1)
    const record = manager.detachTab(id)!
    const newId = manager.attachTab(record)
    expect(fullScreenRegistrations()).toBe(1)

    // the listener resolves the *current* id: fullscreen after re-attach covers the strip
    view.webContents.listeners.get('enter-html-full-screen')!()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: 0,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT,
    })
    view.webContents.listeners.get('leave-html-full-screen')!()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })
    expect(manager.list()[1].id).toBe(newId)
  })

  it('a re-attached docs tab closes through the docs teardown path like any docs tab', async () => {
    const record = manager.detachTab(manager.openDocsTab('/tmp/a.docx'))!
    const newId = manager.attachTab(record)
    await manager.closeTab(newId)
    expect(teardownDocsRenderer).toHaveBeenCalledWith(record.view.webContents)
    expect(record.view.webContents.close).not.toHaveBeenCalled()
    expect(manager.list()).toHaveLength(1)
  })
})
