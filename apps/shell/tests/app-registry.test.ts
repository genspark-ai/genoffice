import { beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * The AppDescriptor registry (src/main/app-registry.ts): the table
 * app-routing, TabManager, the close guards and the MCP bridge all dispatch
 * through. The six app main modules are faked (the registry's faces wrap
 * them); only the table's own behavior — routing, dispatch, menus, agent
 * surfaces — is asserted.
 */

const buildDocsMenu = vi.fn()
const createDocsView = vi.fn(() => ({ webContents: { id: 1 } }))
const docsQueryDirty = vi.fn(() => Promise.resolve(false))
const markDocsNewBlank = vi.fn()
const queueDocsAiContent = vi.fn()
const requestDocsClose = vi.fn(() => Promise.resolve(true))
const setActiveDocsResolver = vi.fn()
const teardownDocsRenderer = vi.fn()

vi.mock('../../docs/src/main/docs-main', () => ({
  buildDocsMenu: (...args: unknown[]) => buildDocsMenu(...(args as [])),
  createDocsView: (...args: unknown[]) => createDocsView(...(args as [])),
  docsQueryDirty: (...args: unknown[]) => docsQueryDirty(...(args as [])),
  markDocsNewBlank: (...args: unknown[]) => markDocsNewBlank(...args),
  queueDocsAiContent: (...args: unknown[]) => queueDocsAiContent(...args),
  requestDocsClose: (...args: unknown[]) => requestDocsClose(...(args as [])),
  setActiveDocsResolver: (...args: unknown[]) => setActiveDocsResolver(...args),
  teardownDocsRenderer: (...args: unknown[]) => teardownDocsRenderer(...args),
}))

const createSheetsView = vi.fn(() => ({ webContents: { id: 2 } }))
const installSheetsMenu = vi.fn()
const nudgeQueuedWorkbook = vi.fn()
const queueWorkbookForView = vi.fn()
const requestSheetsClose = vi.fn(() => Promise.resolve(true))
const setActiveSheetsWebContents = vi.fn()
const setSheetsNewBlank = vi.fn()
const sheetsPendingEditCount = vi.fn(() => 0)

vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsView: (...args: unknown[]) => createSheetsView(...(args as [])),
  installSheetsMenu: (...args: unknown[]) => installSheetsMenu(...args),
  nudgeQueuedWorkbook: (...args: unknown[]) => nudgeQueuedWorkbook(...args),
  queueWorkbookForView: (...args: unknown[]) => queueWorkbookForView(...args),
  requestSheetsClose: (...args: unknown[]) => requestSheetsClose(...(args as [])),
  setActiveSheetsWebContents: (...args: unknown[]) => setActiveSheetsWebContents(...args),
  setSheetsNewBlank: (...args: unknown[]) => setSheetsNewBlank(...args),
  sheetsPendingEditCount: (...args: unknown[]) => sheetsPendingEditCount(...(args as [])),
}))

const createSlidesView = vi.fn(() => ({ webContents: { id: 3 } }))
const discardSlidesRecovery = vi.fn()
const installSlidesMenu = vi.fn()
const requestSlidesClose = vi.fn(() => Promise.resolve(true))
const setActiveSlidesWebContents = vi.fn()
const slidesIsDirty = vi.fn(() => false)

vi.mock('../../slides/src/main/slides-main', () => ({
  createSlidesView: (...args: unknown[]) => createSlidesView(...(args as [])),
  discardSlidesRecovery: (...args: unknown[]) => discardSlidesRecovery(...args),
  installSlidesMenu: (...args: unknown[]) => installSlidesMenu(...args),
  requestSlidesClose: (...args: unknown[]) => requestSlidesClose(...(args as [])),
  setActiveSlidesWebContents: (...args: unknown[]) => setActiveSlidesWebContents(...args),
  slidesIsDirty: (...args: unknown[]) => slidesIsDirty(...(args as [])),
}))

const clearPdfDirty = vi.fn()
const createPdfView = vi.fn(() => ({ webContents: { id: 4 } }))
const pdfIsDirty = vi.fn(() => false)
const requestPdfClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  clearPdfDirty: (...args: unknown[]) => clearPdfDirty(...args),
  createPdfView: (...args: unknown[]) => createPdfView(...(args as [])),
  pdfIsDirty: (...args: unknown[]) => pdfIsDirty(...(args as [])),
  requestPdfClose: (...args: unknown[]) => requestPdfClose(...(args as [])),
}))

const createMarkdownView = vi.fn(() => ({ webContents: { id: 5 } }))
const markdownDiscardPendingAssets = vi.fn(async () => undefined)
const markdownIsDirty = vi.fn(() => false)
const markdownReadText = vi.fn(async () => '')
const markdownSaveToPath = vi.fn(async () => undefined)
const requestMarkdownClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../markdown/src/main/markdown-main', () => ({
  createMarkdownView: (...args: unknown[]) => createMarkdownView(...(args as [])),
  markdownDiscardPendingAssets: (...args: unknown[]) =>
    markdownDiscardPendingAssets(...(args as [])),
  markdownIsDirty: (...args: unknown[]) => markdownIsDirty(...(args as [])),
  markdownReadText: (...args: unknown[]) => markdownReadText(...(args as [])),
  markdownSaveToPath: (...args: unknown[]) => markdownSaveToPath(...(args as [])),
  requestMarkdownClose: (...args: unknown[]) => requestMarkdownClose(...(args as [])),
}))

const createHtmlPresentView = vi.fn(() => ({ webContents: { id: 6 } }))
const createHtmlView = vi.fn(() => ({ webContents: { id: 7 } }))
const htmlDiscardPendingAssets = vi.fn(async () => undefined)
const htmlIsDirty = vi.fn(() => false)
const htmlReadText = vi.fn(async () => '')
const htmlSaveToPath = vi.fn(async () => undefined)
const requestHtmlClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../html/src/main/html-main', () => ({
  createHtmlPresentView: (...args: unknown[]) => createHtmlPresentView(...(args as [])),
  createHtmlView: (...args: unknown[]) => createHtmlView(...(args as [])),
  htmlDiscardPendingAssets: (...args: unknown[]) => htmlDiscardPendingAssets(...(args as [])),
  htmlIsDirty: (...args: unknown[]) => htmlIsDirty(...(args as [])),
  htmlReadText: (...args: unknown[]) => htmlReadText(...(args as [])),
  htmlSaveToPath: (...args: unknown[]) => htmlSaveToPath(...(args as [])),
  requestHtmlClose: (...args: unknown[]) => requestHtmlClose(...(args as [])),
}))

import {
  agentIoFor,
  appForKind,
  appForPath,
  appIsDirty,
  clearShellMenuBuilders,
  documentApps,
  installAppMenu,
  registryAgentIo,
  setShellMenuBuilders,
} from '../src/main/app-registry'
import { LAUNCHABLE_EXTS } from '../src/main/launch-paths'
import type { DocumentTabKind, TabKind } from '../src/shared/tabs-api'
import type { WebContents, WebContentsView } from 'electron'

const DOC_TAB_KINDS: DocumentTabKind[] = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html']

/** a stand-in WebContents for the descriptor faces (they only key or pass it on) */
function wc(id: number): WebContents {
  return { id, isDestroyed: () => false } as unknown as WebContents
}

beforeEach(() => {
  vi.clearAllMocks()
  clearShellMenuBuilders()
})

describe('the descriptor table itself', () => {
  it('registers exactly one app per document tab kind', () => {
    expect(documentApps.map((app) => app.id).sort()).toEqual([...DOC_TAB_KINDS].sort())
    for (const kind of DOC_TAB_KINDS) expect(appForKind(kind)).toBeDefined()
  })

  it('leaves the non-document tab kinds to the shell (home, help)', () => {
    expect(appForKind('home')).toBeUndefined()
    expect(appForKind('help')).toBeUndefined()
  })

  it('routes extensions to exactly one app (no collisions)', () => {
    const seen = new Map<string, DocumentTabKind>()
    for (const app of documentApps) {
      for (const ext of app.extensions) {
        expect(seen.has(ext), `.${ext} is claimed by two apps`).toBe(false)
        seen.set(ext, app.id)
      }
    }
    // the pre-registry routing table had 20 extensions; shrinking it would
    // silently un-route files the OS still offers GenOffice for
    expect(seen.size).toBe(20)
  })

  it('every descriptor carries the faces the shell dispatches through', () => {
    for (const app of documentApps) {
      expect(typeof app.createView).toBe('function')
      expect(typeof app.requestClose).toBe('function')
      // dirty answer: sync or a renderer round-trip, never neither
      expect(Boolean(app.isDirty) || Boolean(app.queryDirty)).toBe(true)
      expect(app.untitledTitle.length).toBeGreaterThan(0)
      expect(app.untitledTitleKey).toMatch(/^untitled[A-Z]/)
    }
  })

  it('advertises every extension the packaged app can launch (and vice versa)', () => {
    const routed = new Set(documentApps.flatMap((app) => app.extensions))
    const launchable = new Set(LAUNCHABLE_EXTS as readonly string[])
    expect([...routed].filter((ext) => !launchable.has(ext))).toEqual([])
    expect([...launchable].filter((ext) => !routed.has(ext))).toEqual([])
  })
})

describe('appForPath (file → app routing)', () => {
  it('routes every extension of every app to that app', () => {
    const routed = new Map<string, DocumentTabKind>()
    for (const app of documentApps) for (const ext of app.extensions) routed.set(ext, app.id)
    for (const [ext, kind] of routed) expect(appForPath(`/x/file.${ext}`)).toBe(kind)
  })

  it('falls back to undefined for an extension no app owns or a pathless file', () => {
    expect(appForPath('/x/file.xyz')).toBeUndefined()
    expect(appForPath('/x/file.doc')).toBeUndefined()
    expect(appForPath('/x/noextension')).toBeUndefined()
    expect(appForPath('')).toBeUndefined()
  })

  it('is case-insensitive, because the filesystem is', () => {
    expect(appForPath('/x/NOTES.MD')).toBe('markdown')
    expect(appForPath('/x/DECK.PPTX')).toBe('slides')
  })
})

describe('per-family dirty dispatch', () => {
  it('answers the five sync families from the main process', async () => {
    const cases: [DocumentTabKind, () => number][] = [
      ['sheets', () => sheetsPendingEditCount.mock.calls.length],
      ['slides', () => slidesIsDirty.mock.calls.length],
      ['pdf', () => pdfIsDirty.mock.calls.length],
      ['markdown', () => markdownIsDirty.mock.calls.length],
      ['html', () => htmlIsDirty.mock.calls.length],
    ]
    for (const [kind, calls] of cases) {
      const contents = wc(11)
      await appIsDirty(kind, contents)
      expect(calls()).toBe(1)
    }
    // the sync families never reach the docs renderer
    expect(docsQueryDirty).not.toHaveBeenCalled()
  })

  it('asks the docs renderer (the one async family)', async () => {
    docsQueryDirty.mockResolvedValueOnce(true)
    await expect(appIsDirty('docs', wc(12))).resolves.toBe(true)
    expect(docsQueryDirty).toHaveBeenCalledTimes(1)
  })

  it('keys the id-based dirty answers by webContents id', () => {
    appForKind('pdf')!.isDirty!(wc(42))
    expect(pdfIsDirty).toHaveBeenCalledWith(42)
    appForKind('sheets')!.isDirty!(wc(43))
    expect(sheetsPendingEditCount).toHaveBeenCalledWith(43)
  })

  it('answers false for the non-document kinds', async () => {
    await expect(appIsDirty('home', wc(1))).resolves.toBe(false)
    await expect(appIsDirty('help', wc(1))).resolves.toBe(false)
  })
})

describe('per-family close guard dispatch', () => {
  it('runs the owning app\u2019s requestClose with the webContents and window', async () => {
    const contents = wc(21)
    const win = { isDestroyed: () => false } as never
    const askers = {
      docs: requestDocsClose,
      sheets: requestSheetsClose,
      slides: requestSlidesClose,
      pdf: requestPdfClose,
      markdown: requestMarkdownClose,
      html: requestHtmlClose,
    }
    for (const kind of DOC_TAB_KINDS) {
      await appForKind(kind)!.requestClose(contents, win)
      expect(askers[kind]).toHaveBeenLastCalledWith(contents, win)
    }
  })
})

describe('createView dispatch', () => {
  it('hands the path straight to the app\u2019s view factory', () => {
    appForKind('docs')!.createView('/tmp/a.docx')
    expect(createDocsView).toHaveBeenCalledWith('/tmp/a.docx')
    appForKind('slides')!.createView('/tmp/b.pptx')
    expect(createSlidesView).toHaveBeenCalledWith('/tmp/b.pptx')
    appForKind('pdf')!.createView('/tmp/c.pdf')
    expect(createPdfView).toHaveBeenCalledWith('/tmp/c.pdf')
  })

  it('sheets queues the workbook on the fresh view (multi-open path binding)', () => {
    const view = appForKind('sheets')!.createView('/tmp/budget.xlsx') as WebContentsView
    expect(createSheetsView).toHaveBeenCalledWith({
      includeAiHandlers: false,
      openingWorkbook: true,
    })
    expect(queueWorkbookForView).toHaveBeenCalledWith(view.webContents, '/tmp/budget.xlsx')
    expect(nudgeQueuedWorkbook).not.toHaveBeenCalled()
  })

  it('a blank sheets view opens without the workbook queue', () => {
    appForKind('sheets')!.createView()
    expect(createSheetsView).toHaveBeenCalledWith({
      includeAiHandlers: false,
      openingWorkbook: false,
    })
    expect(queueWorkbookForView).not.toHaveBeenCalled()
  })

  it('docs teardown is the only one beyond plain webContents.close', () => {
    const contents = wc(81)
    appForKind('docs')!.teardown!(contents)
    expect(teardownDocsRenderer).toHaveBeenCalledWith(contents)
    for (const kind of ['sheets', 'slides', 'pdf', 'markdown', 'html'] as const)
      expect(appForKind(kind)!.teardown).toBeUndefined()
  })
})

describe('open hooks', () => {
  it('docs marks a new-blank view and queues AI content after creation', () => {
    const descriptor = appForKind('docs')!
    descriptor.open?.afterCreate?.(9, { newBlank: true })
    expect(markDocsNewBlank).toHaveBeenCalledWith(9)
    descriptor.open?.afterCreate?.(9, { aiContent: { title: 'x' } as never })
    expect(queueDocsAiContent).toHaveBeenCalledWith(9, { title: 'x' })
  })

  it('sheets sets its blank flag before the view exists, and a spare adopts its path with a nudge', () => {
    const descriptor = appForKind('sheets')!
    descriptor.open?.prepare?.({ newBlank: true })
    expect(setSheetsNewBlank).toHaveBeenCalledTimes(1)
    const view = { webContents: { id: 31 } } as never
    descriptor.open?.spare?.adopt?.(view, '/tmp/spare.xlsx')
    expect(queueWorkbookForView).toHaveBeenCalledWith({ id: 31 }, '/tmp/spare.xlsx')
    expect(nudgeQueuedWorkbook).toHaveBeenCalledWith({ id: 31 })
  })
})

describe('menu dispatch', () => {
  it('installs the app-owned menus for docs/sheets/slides', () => {
    expect(installAppMenu('docs')).toBe(true)
    expect(buildDocsMenu).toHaveBeenCalledTimes(1)
    expect(installAppMenu('sheets')).toBe(true)
    expect(installSheetsMenu).toHaveBeenCalledTimes(1)
    expect(installAppMenu('slides')).toBe(true)
    expect(installSlidesMenu).toHaveBeenCalledTimes(1)
  })

  it('installs the shell-registered menus for pdf/markdown/html', () => {
    const pdf = vi.fn()
    const markdown = vi.fn()
    const html = vi.fn()
    setShellMenuBuilders({ pdf, markdown, html })
    expect(installAppMenu('pdf')).toBe(true)
    expect(pdf).toHaveBeenCalledTimes(1)
    expect(installAppMenu('markdown')).toBe(true)
    expect(markdown).toHaveBeenCalledTimes(1)
    expect(installAppMenu('html')).toBe(true)
    expect(html).toHaveBeenCalledTimes(1)
  })

  it('refuses the shell-owned kinds until a builder is registered', () => {
    expect(installAppMenu('pdf')).toBe(false)
    setShellMenuBuilders({ pdf: vi.fn() })
    expect(installAppMenu('pdf')).toBe(true)
    // an explicit undefined unregisters again
    setShellMenuBuilders({ pdf: undefined })
    expect(installAppMenu('pdf')).toBe(false)
  })

  it('reports the non-document kinds so the caller falls back to the home menu', () => {
    expect(installAppMenu('home')).toBe(false)
    expect(installAppMenu('help')).toBe(false)
    expect(installAppMenu('nonsense' as TabKind)).toBe(false)
  })
})

describe('active-target claims', () => {
  it('docs always claims (resetting to null when another family is focused)', () => {
    const getter = () => wc(51)
    appForKind('docs')!.claimActiveTarget!(getter)
    expect(setActiveDocsResolver).toHaveBeenCalledWith(getter)
    const nullGetter = () => null
    appForKind('docs')!.claimActiveTarget!(nullGetter)
    expect(setActiveDocsResolver).toHaveBeenLastCalledWith(nullGetter)
  })

  it('sheets/slides set their target only while one of their views is active', () => {
    const live = wc(52)
    appForKind('sheets')!.claimActiveTarget!(() => live)
    expect(setActiveSheetsWebContents).toHaveBeenCalledWith(live)
    appForKind('sheets')!.claimActiveTarget!(() => null)
    expect(setActiveSheetsWebContents).toHaveBeenCalledTimes(1)
    appForKind('slides')!.claimActiveTarget!(() => live)
    expect(setActiveSlidesWebContents).toHaveBeenCalledWith(live)
    appForKind('slides')!.claimActiveTarget!(() => null)
    expect(setActiveSlidesWebContents).toHaveBeenCalledTimes(1)
  })

  it('the editor-only families keep no process-global target', () => {
    for (const kind of ['pdf', 'markdown', 'html'] as const)
      expect(appForKind(kind)!.claimActiveTarget).toBeUndefined()
  })
})

describe('the MCP agent surfaces', () => {
  it('pdf is read-only: a plain save and a dirty save-close are both refused with their own reasons', () => {
    const io = agentIoFor('pdf')!
    expect(io.read).toBeUndefined()
    expect(io.save).toBeUndefined()
    expect(() => {
      throw new Error(io.saveDenied)
    }).toThrow(/viewer/)
    expect(io.saveCloseDenied!('six.pdf')).toMatch(/"six.pdf" has unsaved changes/)
    expect(io.saveCloseDenied!('six.pdf')).toMatch(/unsaved:"discard"/)
  })

  it('markdown/html expose the apps\u2019 own read/save/discard', () => {
    const md = agentIoFor('markdown')!
    const contents = wc(61)
    void md.read!(contents)
    expect(markdownReadText).toHaveBeenCalledWith(contents)
    void md.save!(contents, '/tmp/notes.md')
    expect(markdownSaveToPath).toHaveBeenCalledWith(contents, '/tmp/notes.md')
    void md.discard!(contents)
    expect(markdownDiscardPendingAssets).toHaveBeenCalledWith(contents)
    expect(md.discardNeedsFilePath).toBe(true)

    const html = agentIoFor('html')!
    void html.read!(contents)
    expect(htmlReadText).toHaveBeenCalledWith(contents)
    void html.save!(contents, '/tmp/page.html')
    expect(htmlSaveToPath).toHaveBeenCalledWith(contents, '/tmp/page.html')
    expect(html.discardNeedsFilePath).toBe(true)
  })

  it('the session families have no static surface (their controls are runtime-built)', () => {
    expect(agentIoFor('docs')).toBeUndefined()
    expect(agentIoFor('sheets')).toBeUndefined()
    expect(agentIoFor('slides')).toBeUndefined()
  })

  it('registryAgentIo wraps the session controls into the uniform surface', async () => {
    const docs = { runCommand: vi.fn(async () => ({ ok: true })) }
    const sheets = { runCommand: vi.fn(async () => ({ ok: true })) }
    const slides = {
      readDeck: vi.fn(async () => ({ slides: [] })),
      saveDeck: vi.fn(async () => ({ path: 'x' })),
    }
    const agents = registryAgentIo({ docs, sheets, slides })
    const contents = wc(71)

    await agents.docs!.read!(contents)
    expect(docs.runCommand).toHaveBeenCalledWith(71, 'read_document', {})
    await agents.docs!.save!(contents, '/tmp/out.docx')
    expect(docs.runCommand).toHaveBeenLastCalledWith(71, 'save_document', {
      path: '/tmp/out.docx',
      overwrite: true,
    })
    await agents.sheets!.save!(contents, '/tmp/out.xlsx')
    expect(sheets.runCommand).toHaveBeenLastCalledWith(71, 'save_sheet', {
      path: '/tmp/out.xlsx',
      overwrite: true,
    })
    await agents.slides!.read!(contents)
    expect(slides.readDeck).toHaveBeenCalledWith(71)
    await agents.slides!.save!(contents, '/tmp/out.pptx')
    expect(slides.saveDeck).toHaveBeenCalledWith(71, '/tmp/out.pptx', true)
    // a slides discard drops the recovery copies
    agents.slides!.discard!(contents)
    expect(discardSlidesRecovery).toHaveBeenCalledWith(contents)
  })

  it('registryAgentIo still carries the static surfaces without controls', () => {
    const agents = registryAgentIo()
    expect(agents.pdf!.saveDenied).toMatch(/viewer/)
    const contents = wc(82)
    agents.markdown!.read!(contents)
    expect(markdownReadText).toHaveBeenCalledWith(contents)
    agents.html!.read!(contents)
    expect(htmlReadText).toHaveBeenCalledWith(contents)
    expect(agents.docs).toBeUndefined()
  })
})
