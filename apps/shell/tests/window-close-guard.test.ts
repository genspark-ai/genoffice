import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * The shell window's unsaved-changes guard
 * (src/main/window-close-guard.ts): the clean fast path, the per-app prompt
 * walk, the in-flight debounce and its release. The handler under test is the
 * real one — it takes the window and the tab manager as arguments, so it runs
 * without Electron; only the six app-mains that own the prompts are faked.
 */

const docsQueryDirty = vi.fn(() => Promise.resolve(false))
const requestDocsClose = vi.fn(() => Promise.resolve(true))
const requestSheetsClose = vi.fn(() => Promise.resolve(true))
const resetSheetsShuttingDown = vi.fn()
const requestSlidesClose = vi.fn(() => Promise.resolve(true))
const requestPdfClose = vi.fn(() => Promise.resolve(true))
const requestMarkdownClose = vi.fn(() => Promise.resolve(true))
const requestHtmlClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../docs/src/main/docs-main', () => ({
  docsQueryDirty: (...args: unknown[]) => docsQueryDirty(...(args as [])),
  requestDocsClose: (...args: unknown[]) => requestDocsClose(...(args as [])),
}))

vi.mock('../../sheets/src/main/sheets-main', () => ({
  requestSheetsClose: (...args: unknown[]) => requestSheetsClose(...(args as [])),
  resetSheetsShuttingDown: (...args: unknown[]) => resetSheetsShuttingDown(...(args as [])),
}))

vi.mock('../../slides/src/main/slides-main', () => ({
  requestSlidesClose: (...args: unknown[]) => requestSlidesClose(...(args as [])),
}))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  requestPdfClose: (...args: unknown[]) => requestPdfClose(...(args as [])),
}))

vi.mock('../../markdown/src/main/markdown-main', () => ({
  requestMarkdownClose: (...args: unknown[]) => requestMarkdownClose(...(args as [])),
}))

vi.mock('../../html/src/main/html-main', () => ({
  requestHtmlClose: (...args: unknown[]) => requestHtmlClose(...(args as [])),
}))

interface TabStub {
  id: string
  webContents: { id: number }
}

/**
 * The three members the guard uses off a BrowserWindow, with Electron's close
 * contract intact: a close request is cancelled only by preventDefault, and
 * the programmatic win.close() re-enters the same handler.
 */
class FakeWindow {
  destroyed = false
  closeCalls = 0
  handlers = new Map<string, (...args: never[]) => void>()

  on = (event: string, fn: (...args: never[]) => void): void => {
    this.handlers.set(event, fn)
  }

  isDestroyed = (): boolean => this.destroyed

  close = (): void => {
    this.closeCalls += 1
    this.requestClose()
  }

  destroy = (): void => {
    this.destroyed = true
  }

  /** one close request (⌘Q, the red button, a program): true when cancelled */
  requestClose = (): boolean => {
    let prevented = false
    const event = {
      preventDefault: (): void => {
        prevented = true
      },
    }
    this.handlers.get('close')?.(event as never)
    if (!prevented) this.destroy()
    return prevented
  }
}

type DirtyKinds = 'sheets' | 'pdf' | 'markdown' | 'html' | 'slides' | 'docs'

function fakeManager(dirty: Partial<Record<DirtyKinds, TabStub[]>> = {}) {
  return {
    activateTab: vi.fn(),
    dirtySheetsTabs: () => dirty.sheets ?? [],
    dirtyPdfTabs: () => dirty.pdf ?? [],
    dirtyMarkdownTabs: () => dirty.markdown ?? [],
    dirtyHtmlTabs: () => dirty.html ?? [],
    dirtySlidesTabs: () => dirty.slides ?? [],
    docsTabs: () => dirty.docs ?? [],
  }
}

const flush = async (): Promise<void> => {
  for (let i = 0; i < 3; i++) await new Promise((resolve) => setTimeout(resolve, 0))
}

let install: typeof import('../src/main/window-close-guard').installShellCloseGuard
let consoleError: ReturnType<typeof vi.spyOn>

beforeEach(async () => {
  vi.clearAllMocks()
  docsQueryDirty.mockImplementation(() => Promise.resolve(false))
  for (const ask of [
    requestDocsClose,
    requestSheetsClose,
    requestSlidesClose,
    requestPdfClose,
    requestMarkdownClose,
    requestHtmlClose,
  ]) {
    ask.mockImplementation(() => Promise.resolve(true))
  }
  consoleError = vi.spyOn(console, 'error').mockImplementation(() => {})
  install = (await import('../src/main/window-close-guard')).installShellCloseGuard
})

afterEach(() => {
  consoleError.mockRestore()
})

describe('clean fast path', () => {
  it('lets an all-clean close through without cancelling it', () => {
    const win = new FakeWindow()
    install(win as never, fakeManager() as never)
    // nothing dirty: the close event must not be cancelled — ⌘Q is still in
    // flight, and cancelling it would leave the app running with no window
    expect(win.requestClose()).toBe(false)
    expect(win.destroyed).toBe(true)
    // it closed through the original event, not through a re-issued win.close()
    expect(win.closeCalls).toBe(0)
  })

  it('a live docs tab forces the prompt path (dirtiness is renderer-side)', async () => {
    const win = new FakeWindow()
    const docsTab = { id: 'docs-1', webContents: { id: 1 } }
    install(win as never, fakeManager({ docs: [docsTab] }) as never)
    expect(win.requestClose()).toBe(true)
    await flush()
    expect(docsQueryDirty).toHaveBeenCalledWith(docsTab.webContents)
    // a clean docs tab passes through without a prompt and without activation
    expect(requestDocsClose).not.toHaveBeenCalled()
    expect(win.destroyed).toBe(true)
  })
})

describe('dirty close', () => {
  it('cancels the close and asks the owning family, keeping the window on cancel', async () => {
    requestSheetsClose.mockResolvedValueOnce(false)
    const win = new FakeWindow()
    const tab = { id: 'sheets-1', webContents: { id: 7 } }
    const manager = fakeManager({ sheets: [tab] })
    install(win as never, manager as never)
    expect(win.requestClose()).toBe(true)
    await flush()
    expect(requestSheetsClose).toHaveBeenCalledWith(tab.webContents, win)
    // a denied close vetoes the quit that was in flight, so the sheets guard
    // has to prompt again on the next close
    expect(resetSheetsShuttingDown).toHaveBeenCalled()
    expect(win.destroyed).toBe(false)
  })

  it('closes the window once every dirty family has confirmed', async () => {
    const win = new FakeWindow()
    const manager = fakeManager({
      pdf: [{ id: 'p', webContents: { id: 1 } }],
      markdown: [{ id: 'm', webContents: { id: 2 } }],
      html: [{ id: 'h', webContents: { id: 3 } }],
      slides: [{ id: 's', webContents: { id: 4 } }],
    })
    install(win as never, manager as never)
    expect(win.requestClose()).toBe(true)
    await flush()
    for (const ask of [
      requestPdfClose,
      requestMarkdownClose,
      requestHtmlClose,
      requestSlidesClose,
    ]) {
      expect(ask).toHaveBeenCalledWith(expect.anything(), win)
    }
    expect(win.destroyed).toBe(true)
    expect(win.closeCalls).toBe(1)
  })
})

describe('in-flight debounce', () => {
  it('a second close while the prompt is open is cancelled and never re-prompts', async () => {
    let answer: (value: boolean) => void = () => {}
    requestSheetsClose.mockReturnValue(
      new Promise<boolean>((resolve) => {
        answer = resolve
      }),
    )
    const win = new FakeWindow()
    const tab = { id: 'sheets-1', webContents: { id: 7 } }
    install(win as never, fakeManager({ sheets: [tab] }) as never)
    expect(win.requestClose()).toBe(true)
    // the prompt dialog is still open: the retry must not stack on top of it
    expect(win.requestClose()).toBe(true)
    await flush()
    expect(requestSheetsClose).toHaveBeenCalledTimes(1)
    expect(win.destroyed).toBe(false)
    answer(false)
    await flush()
  })
})

describe('failure in the prompt chain', () => {
  it('releases the in-flight flag so the window stays closable', async () => {
    requestSheetsClose.mockRejectedValueOnce(new Error('renderer went away'))
    const win = new FakeWindow()
    const manager = fakeManager({ sheets: [{ id: 'sheets-1', webContents: { id: 7 } }] })
    install(win as never, manager as never)
    expect(win.requestClose()).toBe(true)
    await flush()
    expect(consoleError).toHaveBeenCalled()
    expect(win.destroyed).toBe(false)
    expect(resetSheetsShuttingDown).toHaveBeenCalled()
    // the flag was released, so a later close prompts again instead of being
    // swallowed by the debounce
    expect(win.requestClose()).toBe(true)
    await flush()
    expect(requestSheetsClose).toHaveBeenCalledTimes(2)
  })
})
