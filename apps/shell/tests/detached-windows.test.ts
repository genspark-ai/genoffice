import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * Detached editor windows (src/main/detached-windows.ts): the open-documents
 * listing, the per-kind unsaved-changes guard, tear-off (a window steered by
 * the shell's held pointer) and docking a natively dragged window back onto
 * the shell strip. Electron and the app mains are faked.
 */

const instances: FakeWindow[] = []

class FakeWindow {
  destroyed = false
  title: string
  shown = false
  position: [number, number]
  handlers = new Map<string, (...args: never[]) => void>()
  contentView = { addChildView: vi.fn(), removeChildView: vi.fn() }
  /** the constructor options, for the pinned-widget window shape assertions */
  opts: Record<string, unknown>
  /** the window this one is a child of (pinned widgets) */
  parent: FakeWindow | null = null

  constructor(opts: Record<string, unknown> & { title?: string; show?: boolean }) {
    this.opts = opts
    this.title = (opts.title as string) ?? ''
    this.shown = opts.show !== false
    this.position = [(opts.x as number) ?? 0, (opts.y as number) ?? 0]
    this.parent = (opts.parent as FakeWindow | null) ?? null
    instances.push(this)
  }

  on = (event: string, fn: (...args: never[]) => void): void => {
    this.handlers.set(event, fn)
  }

  emit(event: string, ...args: unknown[]): void {
    this.handlers.get(event)?.(...(args as never[]))
  }

  isDestroyed = (): boolean => this.destroyed
  isMinimized = (): boolean => false
  isVisible = (): boolean => this.shown && !this.destroyed
  restore = vi.fn()
  show = vi.fn(() => {
    this.shown = true
  })
  showInactive = vi.fn(() => {
    this.shown = true
  })
  hide = vi.fn(() => {
    this.shown = false
  })
  focus = vi.fn()
  setParentWindow = vi.fn((parent: FakeWindow | null) => {
    this.parent = parent
  })
  getParentWindow = (): FakeWindow | null => this.parent
  getContentBounds = (): { x: number; y: number; width: number; height: number } => ({
    x: this.position[0],
    y: this.position[1] + 28,
    width: 100,
    height: 100,
  })
  getPosition = (): [number, number] => this.position
  setPosition = vi.fn((x: number, y: number) => {
    this.position = [x, y]
  })
  setTitle = (title: string): void => {
    this.title = title
  }

  getTitle = (): string => {
    if (this.destroyed) throw new Error('window destroyed')
    return this.title
  }

  focused = false
  isFocused = (): boolean => this.focused
  destroy = (): void => {
    this.destroyed = true
    this.emit('closed')
  }
}

const cursor = { x: 0, y: 0 }
const workArea = { x: 0, y: 0, width: 2000, height: 1200 }

/** the pinned widgets' drag strip: a tiny webContents view with a data-URL page */
class FakeStripView {
  webContents = {
    loadURL: vi.fn(),
    setWindowOpenHandler: vi.fn(),
    isDestroyed: () => false,
    close: vi.fn(),
    executeJavaScript: vi.fn(async () => undefined),
  }
  setBounds = vi.fn()
  setVisible = vi.fn()
}

const ipcHandles = new Map<string, (...args: never[]) => unknown>()

/** the pin-cap refusal's dialog, mocked with the electron module */
const { showMessageBox } = vi.hoisted(() => ({ showMessageBox: vi.fn(async () => ({})) }))

vi.mock('electron', () => ({
  app: { on: vi.fn(), quit: vi.fn() },
  BrowserWindow: FakeWindow,
  WebContentsView: FakeStripView,
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      ipcHandles.set(channel, handler),
  },
  screen: {
    getCursorScreenPoint: () => ({ ...cursor }),
    getDisplayNearestPoint: () => ({ workArea }),
  },
  dialog: {
    showMessageBox,
  },
}))

vi.mock('../src/main/desktop-ipc', () => ({
  // only the pin cap the gate reads; the store plumbing stays index.ts's job
  MAX_PINS: 8,
}))

vi.mock('../../docs/src/main/docs-main', () => ({
  // buildDocsMenu/createDocsView/markDocsNewBlank/queueDocsAiContent are
  // referenced by the app registry at module scope; these tests never call them
  buildDocsMenu: vi.fn(),
  createDocsView: vi.fn(),
  markDocsNewBlank: vi.fn(),
  queueDocsAiContent: vi.fn(),
  docsQueryDirty: vi.fn(() => Promise.resolve(false)),
  requestDocsClose: vi.fn(() => Promise.resolve(true)),
  setActiveDocsResolver: vi.fn(),
  teardownDocsRenderer: vi.fn(),
}))

vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsView: vi.fn(),
  installSheetsMenu: vi.fn(),
  nudgeQueuedWorkbook: vi.fn(),
  queueWorkbookForView: vi.fn(),
  setSheetsNewBlank: vi.fn(),
  requestSheetsClose: vi.fn(() => Promise.resolve(true)),
  setActiveSheetsWebContents: vi.fn(),
  sheetsPendingEditCount: vi.fn(() => 0),
  resetSheetsShuttingDown: vi.fn(),
}))

vi.mock('../../slides/src/main/slides-main', () => ({
  createSlidesView: vi.fn(),
  installSlidesMenu: vi.fn(),
  discardSlidesRecovery: vi.fn(),
  requestSlidesClose: vi.fn(() => Promise.resolve(true)),
  setActiveSlidesWebContents: vi.fn(),
  slidesIsDirty: vi.fn(() => false),
}))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  createPdfView: vi.fn(),
  clearPdfDirty: vi.fn(),
  pdfIsDirty: vi.fn(() => false),
  requestPdfClose: vi.fn(() => Promise.resolve(true)),
}))

vi.mock('../../markdown/src/main/markdown-main', () => ({
  createMarkdownView: vi.fn(),
  markdownDiscardPendingAssets: vi.fn(async () => undefined),
  markdownReadText: vi.fn(async () => ''),
  markdownSaveToPath: vi.fn(async () => undefined),
  markdownIsDirty: vi.fn(() => false),
  requestMarkdownClose: vi.fn(() => Promise.resolve(true)),
}))

vi.mock('../../html/src/main/html-main', () => ({
  createHtmlView: vi.fn(),
  createHtmlPresentView: vi.fn(),
  htmlDiscardPendingAssets: vi.fn(async () => undefined),
  htmlReadText: vi.fn(async () => ''),
  htmlSaveToPath: vi.fn(async () => undefined),
  htmlIsDirty: vi.fn(() => false),
  requestHtmlClose: vi.fn(() => Promise.resolve(true)),
}))

vi.mock('../src/main/tab-manager', () => ({
  canonicalPath: (p: string) => p.toLowerCase(),
  // the palette summon hook's own behavior lives in tab-manager.test.ts; here
  // it only needs to install and dispose quietly against the fake view
  installPaletteSummon: () => () => {},
  shellTabManager: () => null,
}))

type DetachedModule = typeof import('../src/main/detached-windows')

let detached: DetachedModule

function fakeView(id: number) {
  return {
    webContents: {
      id,
      isDestroyed: () => false,
      focus: vi.fn(),
      close: vi.fn(),
    },
    setBounds: vi.fn(),
    setVisible: vi.fn(),
  }
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

beforeEach(async () => {
  vi.resetModules()
  instances.length = 0
  cursor.x = 0
  cursor.y = 0
  detached = await import('../src/main/detached-windows')
})

afterEach(() => {
  vi.useRealTimers()
})

describe('detachedOpenDocuments', () => {
  it('lists live windows', async () => {
    detached.createDetachedEditorWindow({
      view: fakeView(22) as never,
      kind: 'sheets',
      title: 'b',
      filePath: 'C:/docs/b.txt',
      applyMenuFor: () => {},
    })
    const docs = await detached.detachedOpenDocuments()
    expect(docs.map((d) => d.id)).toEqual(['detached:22'])
  })

  it('skips windows destroyed during the async dirty check', async () => {
    detached.createDetachedEditorWindow({
      view: fakeView(31) as never,
      kind: 'docs',
      title: 'c',
      filePath: 'C:/docs/c.txt',
      applyMenuFor: () => {},
    })
    const win = instances[0]!
    const { docsQueryDirty } = await import('../../docs/src/main/docs-main')
    vi.mocked(docsQueryDirty).mockImplementationOnce(async () => {
      win.destroyed = true
      return false
    })
    await expect(detached.detachedOpenDocuments()).resolves.toEqual([])
  })

  it('reports dirtiness for every kind through its own family', async () => {
    const { pdfIsDirty } = await import('../../pdf/src/main/pdf-main')
    const { slidesIsDirty } = await import('../../slides/src/main/slides-main')
    vi.mocked(pdfIsDirty).mockReturnValue(true)
    vi.mocked(slidesIsDirty).mockReturnValue(false)
    detached.createDetachedEditorWindow({
      view: fakeView(1) as never,
      kind: 'pdf',
      title: 'scan.pdf',
      applyMenuFor: () => {},
    })
    detached.createDetachedEditorWindow({
      view: fakeView(2) as never,
      kind: 'slides',
      title: 'deck.pptx',
      applyMenuFor: () => {},
    })
    const docs = await detached.detachedOpenDocuments()
    expect(docs.map((d) => [d.kind, d.dirty])).toEqual([
      ['pdf', true],
      ['slides', false],
    ])
  })
})

describe('close guard', () => {
  it('closes a clean window of any kind without a prompt and releases the renderer', async () => {
    const view = fakeView(5)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'markdown',
      title: 'notes.md',
      applyMenuFor: () => {},
    })
    const win = instances[0]!
    const event = { preventDefault: vi.fn() }
    win.emit('close', event)
    expect(event.preventDefault).toHaveBeenCalled()
    await flush()
    const { requestMarkdownClose } = await import('../../markdown/src/main/markdown-main')
    expect(requestMarkdownClose).not.toHaveBeenCalled()
    expect(win.destroyed).toBe(true)
    expect(win.contentView.removeChildView).toHaveBeenCalled()
    expect(view.webContents.close).toHaveBeenCalled()
  })

  it('asks the owning family when the document is dirty and keeps the window on cancel', async () => {
    const { htmlIsDirty, requestHtmlClose } = await import('../../html/src/main/html-main')
    vi.mocked(htmlIsDirty).mockReturnValue(true)
    vi.mocked(requestHtmlClose).mockResolvedValueOnce(false)
    const view = fakeView(6)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'html',
      title: 'page.html',
      applyMenuFor: () => {},
    })
    const win = instances[0]!
    win.emit('close', { preventDefault: vi.fn() })
    await flush()
    expect(requestHtmlClose).toHaveBeenCalledWith(view.webContents, win)
    expect(win.destroyed).toBe(false)
    expect(view.webContents.close).not.toHaveBeenCalled()

    vi.mocked(requestHtmlClose).mockResolvedValueOnce(true)
    win.emit('close', { preventDefault: vi.fn() })
    await flush()
    expect(win.destroyed).toBe(true)
  })

  it('docs never has its webContents closed: the teardown path owns it', async () => {
    const { teardownDocsRenderer } = await import('../../docs/src/main/docs-main')
    const view = fakeView(7)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'docs',
      title: 'a.docx',
      applyMenuFor: () => {},
    })
    expect(detached.closeDetachedWithoutPrompt('detached:7')).toBe(true)
    expect(teardownDocsRenderer).toHaveBeenCalledWith(view.webContents)
    expect(view.webContents.close).not.toHaveBeenCalled()
    expect(detached.isDetachedTabId('detached:7')).toBe(false)
  })
})

describe('focused window as menu target', () => {
  it('resolves the focused detached document with its kind and file', () => {
    detached.createDetachedEditorWindow({
      view: fakeView(8) as never,
      kind: 'pdf',
      title: 'scan.pdf',
      filePath: '/tmp/scan.pdf',
      applyMenuFor: () => {},
    })
    expect(detached.focusedDetachedTab()).toBeUndefined()
    instances[0]!.focused = true
    expect(detached.focusedDetachedTab()).toMatchObject({
      id: 'detached:8',
      kind: 'pdf',
      filePath: '/tmp/scan.pdf',
    })
    expect(detached.focusedDetachedKind()).toBe('pdf')
  })

  it('claims the slides menu target on focus', async () => {
    const { setActiveSlidesWebContents } = await import('../../slides/src/main/slides-main')
    const view = fakeView(9)
    const applyMenuFor = vi.fn()
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'slides',
      title: 'deck.pptx',
      applyMenuFor,
    })
    instances[0]!.emit('focus')
    expect(setActiveSlidesWebContents).toHaveBeenCalledWith(view.webContents)
    expect(applyMenuFor).toHaveBeenCalledWith('slides')
  })
})

describe('tear-off', () => {
  it('creates the window under the pointer without focus and steers it with the pointer', () => {
    const view = fakeView(10)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'sheets',
      title: 'x.xlsx',
      applyMenuFor: () => {},
      tearOffAt: { x: 500, y: 300 },
    })
    const win = instances[0]!
    expect(win.showInactive).toHaveBeenCalled()
    expect(win.focus).not.toHaveBeenCalled()
    expect(view.webContents.focus).not.toHaveBeenCalled()
    // the pointer lands inside the title bar: a bit right of and below the origin
    const [x0, y0] = win.position
    expect(x0).toBeLessThan(500)
    expect(y0).toBeLessThan(300)
    expect(detached.isTearingOff()).toBe(true)

    detached.dragTornWindow(560, 340)
    expect(win.setPosition).toHaveBeenLastCalledWith(x0 + 60, y0 + 40)

    // steered to the screen edge: a (-0.5, 0) offset must not reach Electron as -0
    detached.dragTornWindow(500 - x0 - 0.25, 300 - y0 - 0.25)
    const [ex, ey] = win.setPosition.mock.lastCall as [number, number]
    expect(Object.is(ex, -0)).toBe(false)
    expect(Object.is(ey, -0)).toBe(false)
    expect([ex, ey]).toEqual([0, 0])

    detached.endTornDrag()
    expect(win.focus).toHaveBeenCalled()
    expect(view.webContents.focus).toHaveBeenCalled()
    expect(detached.isTearingOff()).toBe(false)
    // an ordinary detached window now: steering is over
    win.setPosition.mockClear()
    detached.dragTornWindow(600, 400)
    expect(win.setPosition).not.toHaveBeenCalled()
  })

  it('keeps the window on the pointer display', () => {
    detached.createDetachedEditorWindow({
      view: fakeView(11) as never,
      kind: 'docs',
      title: 'a.docx',
      applyMenuFor: () => {},
      tearOffAt: { x: 10, y: 5 },
    })
    expect(instances[0]!.position).toEqual([workArea.x, workArea.y])
  })

  it('takeTornTab hands the live view back and destroys the window without a prompt', () => {
    const view = fakeView(12)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'pdf',
      title: 'scan.pdf',
      filePath: '/tmp/scan.pdf',
      applyMenuFor: () => {},
      tearOffAt: { x: 500, y: 300 },
    })
    const win = instances[0]!
    const tab = detached.takeTornTab()
    expect(tab).toMatchObject({ kind: 'pdf', title: 'scan.pdf', filePath: '/tmp/scan.pdf' })
    expect(tab!.view).toBe(view)
    expect(win.contentView.removeChildView).toHaveBeenCalledWith(view)
    expect(win.destroyed).toBe(true)
    expect(view.webContents.close).not.toHaveBeenCalled()
    expect(detached.isDetachedTabId('detached:12')).toBe(false)
    expect(detached.takeTornTab()).toBeNull()
  })

  it('a torn-off window does not dock through the native-drag path', () => {
    const host = {
      band: vi.fn(() => ({ x: 0, y: 0, width: 2000, height: 56 })),
      preview: vi.fn(),
      dock: vi.fn(),
    }
    detached.setDockHost(host)
    detached.createDetachedEditorWindow({
      view: fakeView(13) as never,
      kind: 'docs',
      title: 'a.docx',
      applyMenuFor: () => {},
      tearOffAt: { x: 500, y: 20 },
    })
    cursor.x = 500
    cursor.y = 20
    instances[0]!.emit('move')
    expect(host.preview).not.toHaveBeenCalled()
  })
})

describe('docking a natively dragged window', () => {
  const band = { x: 100, y: 200 - 8, width: 1000, height: 56 }

  function host() {
    return { band: vi.fn(() => band), preview: vi.fn(), dock: vi.fn() }
  }

  it('previews the slot while the cursor is over the strip and docks after the dwell', () => {
    vi.useFakeTimers()
    const h = host()
    detached.setDockHost(h)
    const view = fakeView(14)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'slides',
      title: 'deck.pptx',
      filePath: '/tmp/deck.pptx',
      applyMenuFor: () => {},
    })
    const win = instances[0]!
    cursor.x = 400
    cursor.y = 210
    win.emit('move')
    expect(h.preview).toHaveBeenLastCalledWith(300)
    expect(h.dock).not.toHaveBeenCalled()

    vi.advanceTimersByTime(500)
    expect(h.dock).toHaveBeenCalledTimes(1)
    expect(h.dock.mock.calls[0][0]).toMatchObject({
      kind: 'slides',
      title: 'deck.pptx',
      filePath: '/tmp/deck.pptx',
    })
    expect(h.dock.mock.calls[0][0].view).toBe(view)
    expect(win.destroyed).toBe(true)
    expect(view.webContents.close).not.toHaveBeenCalled()
    expect(detached.detachedFilePaths()).toEqual([])
  })

  it('clears the preview and the pending dock when the cursor leaves the strip', () => {
    vi.useFakeTimers()
    const h = host()
    detached.setDockHost(h)
    detached.createDetachedEditorWindow({
      view: fakeView(15) as never,
      kind: 'docs',
      title: 'a.docx',
      applyMenuFor: () => {},
    })
    const win = instances[0]!
    cursor.x = 400
    cursor.y = 210
    win.emit('move')
    expect(h.preview).toHaveBeenLastCalledWith(300)
    cursor.y = 600
    win.emit('move')
    expect(h.preview).toHaveBeenLastCalledWith(null)
    vi.advanceTimersByTime(1000)
    expect(h.dock).not.toHaveBeenCalled()
    expect(win.destroyed).toBe(false)
  })

  it('does not dock when the cursor left the strip before the dwell elapsed', () => {
    vi.useFakeTimers()
    const h = host()
    detached.setDockHost(h)
    detached.createDetachedEditorWindow({
      view: fakeView(16) as never,
      kind: 'pdf',
      title: 'scan.pdf',
      applyMenuFor: () => {},
    })
    cursor.x = 400
    cursor.y = 210
    instances[0]!.emit('move')
    // the last move landed inside, but the pointer kept going without another event
    cursor.y = 900
    vi.advanceTimersByTime(1000)
    expect(h.dock).not.toHaveBeenCalled()
    expect(h.preview).toHaveBeenLastCalledWith(null)
  })

  it('ignores moves when there is no shell to dock into', () => {
    vi.useFakeTimers()
    const h = { band: vi.fn(() => null), preview: vi.fn(), dock: vi.fn() }
    detached.setDockHost(h)
    detached.createDetachedEditorWindow({
      view: fakeView(17) as never,
      kind: 'docs',
      title: 'a.docx',
      applyMenuFor: () => {},
    })
    instances[0]!.emit('move')
    vi.advanceTimersByTime(1000)
    expect(h.preview).not.toHaveBeenCalled()
    expect(h.dock).not.toHaveBeenCalled()
  })
})

describe('dockBandFor', () => {
  it('is the strip band of a visible shell window, null otherwise', () => {
    const win = new FakeWindow({ x: 10, y: 20 })
    expect(detached.dockBandFor(win as never)).toEqual({
      x: 10,
      y: 20 + 28 - 8,
      width: 100,
      height: 40 + 16,
    })
    win.shown = false
    expect(detached.dockBandFor(win as never)).toBeNull()
    expect(detached.dockBandFor(null)).toBeNull()
  })
})

// ── pinned desktop widgets ──────────────────────────────

describe('pinned widgets', () => {
  /** a shell window to pin over, at a known spot */
  const shellWindow = (): FakeWindow => new FakeWindow({ title: 'GenOffice', x: 40, y: 20 })

  /** one pinned widget; returns the window, the editor view and the strip */
  function pinWidget(id: number, over: Record<string, unknown> = {}) {
    const view = fakeView(id)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'sheets',
      title: 'budget.xlsx',
      filePath: '/docs/budget.xlsx',
      applyMenuFor: () => {},
      pin: { parent: shellWindow() as never, ...over },
    })
    const win = instances.at(-1)!
    const strip = win.contentView.addChildView.mock.calls
      .map((args) => args[0] as unknown as FakeStripView)
      .find((candidate) => candidate instanceof FakeStripView)
    return { win, view, strip }
  }

  it('creates a frameless child window carrying a drag strip, shown without focus', () => {
    const { win, view, strip } = pinWidget(30)
    expect(win.opts).toMatchObject({
      frame: false,
      resizable: true,
      show: false,
      skipTaskbar: true,
    })
    // a child window: the widget stacks above the desktop it decorates
    expect(win.parent).toBeInstanceOf(FakeWindow)
    // the strip is a second view on top, loading its self-contained page
    expect(strip).toBeTruthy()
    expect(strip!.webContents.loadURL).toHaveBeenCalledTimes(1)
    expect(decodeURIComponent(String(strip!.webContents.loadURL.mock.calls[0]))).toContain(
      'genoffice-pin:dock',
    )
    expect(strip!.webContents.setWindowOpenHandler).toHaveBeenCalledTimes(1)
    // the editor view sits below the strip; nothing steals focus
    expect(view.setBounds).toHaveBeenCalledWith({ x: 0, y: 28, width: 100, height: 72 })
    expect(strip!.setBounds).toHaveBeenCalledWith({ x: 0, y: 0, width: 100, height: 28 })
    expect(win.focus).not.toHaveBeenCalled()
    expect(view.webContents.focus).not.toHaveBeenCalled()
  })

  it('a given bounds lands on the window as-is', () => {
    const { win } = pinWidget(31, {
      bounds: { x: 640, y: 360, width: 480, height: 360 },
    })
    expect(win.opts).toMatchObject({ x: 640, y: 360, width: 480, height: 360 })
  })

  it('closing docks the document back: the live view re-homes, no prompt runs', async () => {
    const host = { band: vi.fn(() => null), preview: vi.fn(), dock: vi.fn() }
    detached.setDockHost(host)
    const { win, view } = pinWidget(32)
    const event = { preventDefault: vi.fn() }
    win.emit('close', event)
    expect(event.preventDefault).toHaveBeenCalled()
    expect(host.dock).toHaveBeenCalledTimes(1)
    expect(host.dock.mock.calls[0][0]).toMatchObject({
      kind: 'sheets',
      title: 'budget.xlsx',
      filePath: '/docs/budget.xlsx',
    })
    expect(host.dock.mock.calls[0][0].view).toBe(view)
    expect(win.destroyed).toBe(true)
    expect(view.webContents.close).not.toHaveBeenCalled()
    const { requestSheetsClose } = await import('../../sheets/src/main/sheets-main')
    expect(requestSheetsClose).not.toHaveBeenCalled()
  })

  it('the strip dock control runs the same dock-back path', () => {
    const host = { band: vi.fn(() => null), preview: vi.fn(), dock: vi.fn() }
    detached.setDockHost(host)
    const { win, strip } = pinWidget(33)
    const handler = strip!.webContents.setWindowOpenHandler.mock.calls[0][0] as ({
      url,
    }: {
      url: string
    }) => { action: 'deny' }
    const result = handler({ url: 'genoffice-pin:dock' })
    expect(result).toEqual({ action: 'deny' })
    expect(host.dock).toHaveBeenCalledTimes(1)
    expect(win.destroyed).toBe(true)
  })

  it('persisting through the injected store: pin writes, a move debounces, dock clears', () => {
    vi.useFakeTimers()
    const writes: unknown[][] = []
    detached.setDesktopPinStore({
      read: () => [],
      write: (pins) => writes.push(pins),
    })
    const { win } = pinWidget(34)
    // the pin landed: one immediate write carrying kind + path + rect (the
    // fake shell sits at (40,20): its content rect is (40,48), the widget's
    // content rect reports (40,76))
    expect(writes).toHaveLength(1)
    expect(writes[0]).toEqual([
      {
        kind: 'sheets',
        filePath: '/docs/budget.xlsx',
        bounds: { x: 40, y: 76, width: 100, height: 100 },
      },
    ])
    // a move re-persists once the 500ms window closes, not per event
    win.emit('moved')
    win.emit('moved')
    win.emit('resized')
    expect(writes).toHaveLength(1)
    vi.advanceTimersByTime(500)
    expect(writes).toHaveLength(2)
    // docking back clears the store
    const host = { band: vi.fn(() => null), preview: vi.fn(), dock: vi.fn() }
    detached.setDockHost(host)
    win.emit('close', { preventDefault: vi.fn() })
    expect(writes.at(-1)).toEqual([])
    detached.setDesktopPinStore(null)
    vi.useRealTimers()
  })

  it('an untitled widget is never persisted (nothing to reopen)', () => {
    const writes: unknown[][] = []
    detached.setDesktopPinStore({ read: () => [], write: (pins) => writes.push(pins) })
    const view = fakeView(35)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'markdown',
      title: 'Untitled Markdown',
      applyMenuFor: () => {},
      pin: { parent: shellWindow() as never },
    })
    // the pin mutation still writes — an empty list, since nothing is restorable
    expect(writes).toEqual([[]])
    // but it is a live widget in the listing
    expect(detached.pinnedDesktopPins()).toMatchObject([{ kind: 'markdown', filePath: null }])
    detached.setDesktopPinStore(null)
  })

  it('the widget list carries the runtime id, the title and the screen rect', () => {
    pinWidget(36)
    expect(detached.pinnedDesktopPins()).toEqual([
      {
        id: 'detached:36',
        kind: 'sheets',
        title: 'budget.xlsx',
        filePath: '/docs/budget.xlsx',
        bounds: { x: 40, y: 76, width: 100, height: 100 },
      },
    ])
  })

  it('visibility follows the desktop view and the host surface, showing inactive', () => {
    detached.registerDetachedPinsIpc()
    const { win } = pinWidget(37)
    const invoke = (channel: string, ...args: unknown[]) =>
      ipcHandles.get(channel)!({ sender: { id: 1 } } as never, ...(args as never[]))
    // the renderer mounted the desktop view
    expect(invoke('home:desktop-pins-set-visible', true)).toBe(true)
    expect(win.showInactive).toHaveBeenCalled()
    expect(win.hide).not.toHaveBeenCalled()
    // another tab covers the desktop: the widget hides with the surface
    detached.setPinsHostSurfaceActive(false)
    expect(win.hide).toHaveBeenCalled()
    // and comes back only with both signals up
    win.showInactive.mockClear()
    detached.setPinsHostSurfaceActive(true)
    expect(win.showInactive).toHaveBeenCalled()
    // leaving the desktop view hides it again
    expect(invoke('home:desktop-pins-set-visible', false)).toBe(true)
    expect(win.hide).toHaveBeenCalled()
    // junk payloads change nothing
    expect(invoke('home:desktop-pins-set-visible', 'yes')).toBe(false)
  })

  it('the pins IPC serves the list and raises one widget without focus', () => {
    detached.registerDetachedPinsIpc()
    const first = pinWidget(38)
    const second = pinWidget(39)
    const invoke = (channel: string, ...args: unknown[]) =>
      ipcHandles.get(channel)!({ sender: { id: 1 } } as never, ...(args as never[]))
    expect(invoke('home:desktop-pins')).toEqual([
      expect.objectContaining({ id: 'detached:38' }),
      expect.objectContaining({ id: 'detached:39' }),
    ])
    expect(invoke('home:desktop-pins-activate', 'detached:39')).toBe(true)
    expect(second.win.showInactive).toHaveBeenCalled()
    expect(second.win.focus).not.toHaveBeenCalled()
    expect(first.win.showInactive).not.toHaveBeenCalled()
    expect(invoke('home:desktop-pins-activate', 'detached:999')).toBe(false)
    expect(invoke('home:desktop-pins-activate', 42)).toBe(false)
  })

  it('re-parents to a (re)created shell and unparents before one closes', () => {
    const shellA = shellWindow()
    const { win } = pinWidget(40, { parent: shellA })
    expect(win.parent).toBe(shellA)
    // the shell is closing: the widget must not be destroyed with it
    detached.detachPinnedWidgetsFromShell(shellA)
    expect(win.parent).toBeNull()
    // a new shell window: the widget re-parents to keep its stacking
    const shellB = shellWindow()
    detached.attachPinnedWidgetsToShell(shellB)
    expect(win.parent).toBe(shellB)
  })

  it('a document already pinned shows its widget instead of a second copy', () => {
    detached.setDockHost({ band: vi.fn(() => null), preview: vi.fn(), dock: vi.fn() })
    const first = pinWidget(41)
    expect(detached.focusPinnedWidget('/DOCS/BUDGET.XLSX')).toBe(true)
    expect(first.win.showInactive).toHaveBeenCalled()
    // no second window was created for the pinned document
    expect(instances).toHaveLength(2)
    expect(detached.focusPinnedWidget('/docs/other.xlsx')).toBe(false)
  })

  it('refuses the 9th widget: an error dialog, the document in a regular window', () => {
    showMessageBox.mockClear()
    const writes: unknown[][] = []
    detached.setDesktopPinStore({ read: () => [], write: (pins) => writes.push(pins) })
    detached.setDockHost({ band: vi.fn(() => null), preview: vi.fn(), dock: vi.fn() })
    for (let i = 0; i < 8; i++) pinWidget(50 + i)
    expect(detached.pinnedDesktopPins()).toHaveLength(8)
    expect(showMessageBox).not.toHaveBeenCalled()
    // the capped pin: no widget is created, the refusal is said out loud
    const view = fakeView(90)
    detached.createDetachedEditorWindow({
      view: view as never,
      kind: 'sheets',
      title: 'ninth.xlsx',
      filePath: '/docs/ninth.xlsx',
      applyMenuFor: () => {},
      pin: { parent: shellWindow() as never },
    })
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    const [parent, options] = showMessageBox.mock.calls[0] as unknown as [
      FakeWindow,
      { message: string; detail: string },
    ]
    expect(parent).toBeInstanceOf(FakeWindow)
    expect(options.message).toBeTruthy()
    expect(options.detail).toContain('8')
    // the fallback window is an ordinary editor window: framed, unparented, no strip
    const win = instances.at(-1)!
    expect(win.opts).not.toHaveProperty('frame')
    expect(win.parent).toBeNull()
    expect(win.contentView.addChildView).toHaveBeenCalledTimes(1)
    // and it is not a widget: not in the listing, not persisted
    expect(detached.pinnedDesktopPins()).toHaveLength(8)
    expect(detached.pinnedDesktopPins().map((p) => p.id)).not.toContain('detached:90')
    expect(writes).toHaveLength(8)
    expect(writes.at(-1)).toHaveLength(8)
    detached.setDesktopPinStore(null)
  })
})
