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
  shown: boolean
  position: [number, number]
  handlers = new Map<string, (...args: never[]) => void>()
  contentView = { addChildView: vi.fn(), removeChildView: vi.fn() }

  constructor(opts: { title?: string; show?: boolean; x?: number; y?: number }) {
    this.title = opts.title ?? ''
    this.shown = opts.show !== false
    this.position = [opts.x ?? 0, opts.y ?? 0]
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
  focus = vi.fn()
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

vi.mock('electron', () => ({
  BrowserWindow: FakeWindow,
  screen: {
    getCursorScreenPoint: () => ({ ...cursor }),
    getDisplayNearestPoint: () => ({ workArea }),
  },
}))

vi.mock('../../docs/src/main/docs-main', () => ({
  docsQueryDirty: vi.fn(() => Promise.resolve(false)),
  requestDocsClose: vi.fn(() => Promise.resolve(true)),
  setActiveDocsResolver: vi.fn(),
  teardownDocsRenderer: vi.fn(),
}))

vi.mock('../../sheets/src/main/sheets-main', () => ({
  requestSheetsClose: vi.fn(() => Promise.resolve(true)),
  setActiveSheetsWebContents: vi.fn(),
  sheetsPendingEditCount: vi.fn(() => 0),
}))

vi.mock('../../slides/src/main/slides-main', () => ({
  requestSlidesClose: vi.fn(() => Promise.resolve(true)),
  setActiveSlidesWebContents: vi.fn(),
  slidesIsDirty: vi.fn(() => false),
}))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  pdfIsDirty: vi.fn(() => false),
  requestPdfClose: vi.fn(() => Promise.resolve(true)),
}))

vi.mock('../../markdown/src/main/markdown-main', () => ({
  markdownIsDirty: vi.fn(() => false),
  requestMarkdownClose: vi.fn(() => Promise.resolve(true)),
}))

vi.mock('../../html/src/main/html-main', () => ({
  htmlIsDirty: vi.fn(() => false),
  requestHtmlClose: vi.fn(() => Promise.resolve(true)),
}))

vi.mock('../src/main/tab-manager', () => ({
  canonicalPath: (p: string) => p.toLowerCase(),
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
