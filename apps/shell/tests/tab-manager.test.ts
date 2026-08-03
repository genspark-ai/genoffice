import { beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * TabManager (src/main/tab-manager.ts): tab list state, activation,
 * close guards, and view lifecycle inside the shell's single window.
 * Electron and the per-module main entrypoints are mocked; only the
 * manager's own observable behavior is asserted.
 */

interface FakeWebContents {
  id: number
  on: ReturnType<typeof vi.fn>
  close: ReturnType<typeof vi.fn>
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
      close: vi.fn(),
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

vi.mock('../../docs/src/main/docs-main', () => ({
  createDocsView: (...args: unknown[]) => createDocsView(...(args as [])),
  docsQueryDirty: (...args: unknown[]) => docsQueryDirty(...(args as [])),
  markDocsNewBlank: (...args: unknown[]) => markDocsNewBlank(...args),
  requestDocsClose: (...args: unknown[]) => requestDocsClose(...(args as [])),
  setActiveDocsResolver: (...args: unknown[]) => setActiveDocsResolver(...args),
}))

const createPdfView = vi.fn(() => makeFakeView())
const pdfIsDirty = vi.fn(() => false)
const requestPdfClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  createPdfView: (...args: unknown[]) => createPdfView(...(args as [])),
  pdfIsDirty: (...args: unknown[]) => pdfIsDirty(...(args as [])),
  requestPdfClose: (...args: unknown[]) => requestPdfClose(...(args as [])),
}))

const createSheetsView = vi.fn(() => makeFakeView())
const requestSheetsClose = vi.fn(() => Promise.resolve(true))
const setActiveSheetsWebContents = vi.fn()
const setSheetsNewBlank = vi.fn()
const sheetsPendingEditCount = vi.fn(() => 0)

vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsView: (...args: unknown[]) => createSheetsView(...(args as [])),
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

import { TabManager } from '../src/main/tab-manager'

const TAB_STRIP_HEIGHT = 40
const WINDOW_WIDTH = 800
const WINDOW_HEIGHT = 600

interface FakeShellWindow {
  on: ReturnType<typeof vi.fn>
  getContentBounds: () => { x: number; y: number; width: number; height: number }
  contentView: {
    addChildView: ReturnType<typeof vi.fn>
    removeChildView: ReturnType<typeof vi.fn>
  }
}

function makeShellWindow(): FakeShellWindow {
  return {
    on: vi.fn(),
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
