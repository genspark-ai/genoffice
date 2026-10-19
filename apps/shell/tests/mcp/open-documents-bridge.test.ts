import { describe, expect, it, vi } from 'vitest'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { WebContents } from 'electron'
import type { OpenDocumentTab } from '../../src/shared/tabs-api'
import {
  createOpenDocumentsControl,
  createOpenTargetResolver,
} from '../../src/main/mcp/open-documents-bridge'
import type { OpenDocumentsBridgeDeps } from '../../src/main/mcp/open-documents-bridge'
import { createOpenDocumentTools } from '../../src/main/mcp/tools/open-documents-tools'

/**
 * `open_documents` close/read through the shell-side control.
 *
 * These exercise the wiring the tool-level tests stub out: which family writer a
 * close-save reaches, and what a discard is expected to clean up. The point is
 * the per-family branching, which is where the untitled-html crash and the
 * slides recovery leak lived.
 */

function tab(overrides: Partial<OpenDocumentTab> & { id: string }): OpenDocumentTab {
  return { kind: 'markdown', title: 'notes.md', active: false, dirty: true, ...overrides } as never
}

/** minimal stand-in for the WebContents the bridges address */
function contentsFor(): WebContents {
  return { id: 7, isDestroyed: () => false } as unknown as WebContents
}

async function controlWith(
  documents: OpenDocumentTab[],
  overrides: Partial<OpenDocumentsBridgeDeps> = {},
): Promise<{
  call: (args: Record<string, unknown>) => Promise<Record<string, unknown>>
  deps: OpenDocumentsBridgeDeps
}> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-bridge-'))
  const deps: OpenDocumentsBridgeDeps = {
    list: async () => documents,
    webContentsFor: () => contentsFor(),
    closeTab: () => true,
    defaultSaveDir: () => dir,
    ...overrides,
  }
  const control = createOpenDocumentsControl(deps)
  const [tool] = createOpenDocumentTools({ control, defaultSaveDir: () => dir })
  return {
    deps,
    call: async (args) => {
      const result = await tool!.handler(args)
      return result as Record<string, unknown>
    },
  }
}

describe('open_documents close: untitled documents', () => {
  // Regression: an untitled html document threw "family \"html\" has no MCP
  // generation format" because the target extension was asked of the generation
  // registry, which only covers the three headless create_* families.
  it.each([
    ['docs', '.docx'],
    ['sheets', '.xlsx'],
    ['slides', '.pptx'],
    ['markdown', '.md'],
    ['html', '.html'],
  ] as const)('saves an untitled dirty %s tab before closing it', async (kind, ext) => {
    const doc = tab({ id: 't1', kind, title: `untitled ${kind}` })
    const { call } = await controlWith([doc], {
      docs: { openBlankTab: async () => 1, runCommand: async () => ({ ok: true }) },
      sheets: { openBlankTab: async () => 1, runCommand: async () => ({ ok: true }) },
      slides: {
        openBlankTab: async () => 1,
        runTxn: async () => ({}),
        readDeck: async () => ({}),
        saveDeck: async (_wcId, path) => ({ path }),
      },
      markdown: {
        read: async () => '',
        save: async () => undefined,
        discard: async () => undefined,
      },
      html: { read: async () => '', save: async () => undefined, discard: async () => undefined },
    })

    const result = await call({ action: 'close', target: 't1' })
    expect(result.closed).toBe(true)
    expect(String(result.savedTo).endsWith(ext)).toBe(true)
  })
})

describe('open_documents close: discard cleanup', () => {
  // Regression: discard cleaned markdown/html staged assets only. slides keeps
  // its own crash-recovery copy, so a discarded deck was offered back on the
  // next open.
  it('clears a slides document\u2019s recovery copies on discard', async () => {
    const discards: number[] = []
    const doc = tab({ id: 't3', kind: 'slides', title: 'deck.pptx', filePath: 'C:/docs/deck.pptx' })
    const { call } = await controlWith([doc], {
      slidesDiscard: () => discards.push(1),
    })

    const result = await call({ action: 'close', target: 't3', unsaved: 'discard' })
    expect(result.closed).toBe(true)
    expect(result.discardedUnsavedChanges).toBe(true)
    expect(discards).toHaveLength(1)
  })

  it('does not touch the recovery copy when the close saves instead', async () => {
    const discards: number[] = []
    const saved: string[] = []
    const doc = tab({ id: 't4', kind: 'slides', title: 'deck.pptx', filePath: 'C:/docs/deck.pptx' })
    const { call } = await controlWith([doc], {
      slidesDiscard: () => discards.push(1),
      slides: {
        openBlankTab: async () => 1,
        runTxn: async () => ({}),
        readDeck: async () => ({}),
        saveDeck: async (_wcId, path) => {
          saved.push(path)
          return { path }
        },
      },
    })

    const result = await call({ action: 'close', target: 't4' })
    expect(result.savedTo).toBe('C:/docs/deck.pptx')
    expect(saved).toEqual(['C:/docs/deck.pptx'])
    expect(discards).toEqual([])
  })

  it('still releases markdown staged assets on discard', async () => {
    const discard = vi.fn(async () => undefined)
    const doc = tab({ id: 't5', kind: 'markdown', title: 'notes.md', filePath: 'C:/docs/notes.md' })
    const { call } = await controlWith([doc], {
      markdown: { read: async () => '', save: async () => undefined, discard },
    })

    await call({ action: 'close', target: 't5', unsaved: 'discard' })
    expect(discard).toHaveBeenCalledTimes(1)
  })
})

describe('open_documents close: PDF guard', () => {
  it('refuses to close a PDF tab, naming the reason', async () => {
    const doc = tab({ id: 't6', kind: 'pdf', title: 'six.pdf', filePath: 'C:/docs/six.pdf' })
    const { call } = await controlWith([doc])
    await expect(call({ action: 'close', target: 't6' })).rejects.toThrow(/viewer/)
  })
})

describe('open_documents save: writes without closing', () => {
  it('writes a document back to its own file and leaves the tab open', async () => {
    const saved: string[] = []
    const doc = tab({ id: 't1', kind: 'slides', title: 'deck.pptx', filePath: 'C:/docs/deck.pptx' })
    const { call } = await controlWith([doc], {
      slides: {
        openBlankTab: async () => 1,
        runTxn: async () => ({}),
        readDeck: async () => ({}),
        saveDeck: async (_wcId, path) => {
          saved.push(path)
          return { path }
        },
      },
    })

    const result = await call({ action: 'save', target: 't1' })
    expect(result.saved).toBe(true)
    expect(result.savedTo).toBe('C:/docs/deck.pptx')
    expect(result.closed).toBeUndefined()
    expect(saved).toEqual(['C:/docs/deck.pptx'])
  })

  it('refuses to save a PDF, naming the reason', async () => {
    const doc = tab({ id: 't6', kind: 'pdf', title: 'six.pdf', filePath: 'C:/docs/six.pdf' })
    const { call } = await controlWith([doc])
    await expect(call({ action: 'save', target: 't6' })).rejects.toThrow(/viewer/)
  })
})

describe('open target resolver: focusing the document the agent edits', () => {
  it('activates the named tab and reveals the window', async () => {
    const activate = vi.fn()
    const revealWindow = vi.fn()
    const doc = tab({
      id: 't9',
      kind: 'sheets',
      title: 'book.xlsx',
      filePath: 'C:/docs/book.xlsx',
      active: false,
    })
    const resolve = createOpenTargetResolver({
      list: async () => [doc],
      webContentsFor: () => contentsFor(),
      activate,
      revealWindow,
    })

    await expect(resolve('t9', 'xlsx', { focus: true })).resolves.toBe(7)
    expect(activate).toHaveBeenCalledWith('t9')
    expect(revealWindow).toHaveBeenCalledTimes(1)
  })

  it('leaves the UI alone for a read, and for a tab already in front', async () => {
    const activate = vi.fn()
    const revealWindow = vi.fn()
    const back = tab({ id: 't1', kind: 'sheets', title: 'back.xlsx', active: false })
    const front = tab({ id: 't2', kind: 'sheets', title: 'front.xlsx', active: true })
    const resolve = createOpenTargetResolver({
      list: async () => [back, front],
      webContentsFor: () => contentsFor(),
      activate,
      revealWindow,
    })

    // read tools pass no focus option
    await resolve('t1', 'xlsx')
    expect(activate).not.toHaveBeenCalled()

    // an already-active tab needs no switch
    await resolve('t2', 'xlsx', { focus: true })
    expect(activate).not.toHaveBeenCalled()
    expect(revealWindow).not.toHaveBeenCalled()
  })

  it('still resolves the edit when the UI call fails', async () => {
    const doc = tab({ id: 't3', kind: 'xlsx' as never, title: 'x.xlsx', active: false })
    const resolve = createOpenTargetResolver({
      list: async () => [{ ...doc, kind: 'sheets' }],
      webContentsFor: () => contentsFor(),
      activate: () => {
        throw new Error('no such tab')
      },
    })
    // showing the tab is a courtesy: the caller's edit must still go through
    await expect(resolve('t3', 'xlsx', { focus: true })).resolves.toBe(7)
  })

  it('refuses a document of the wrong family', async () => {
    const doc = tab({ id: 't4', kind: 'docs', title: 'doc.docx', active: false })
    const resolve = createOpenTargetResolver({
      list: async () => [doc],
      webContentsFor: () => contentsFor(),
    })
    await expect(resolve('t4', 'xlsx', { focus: true })).rejects.toThrow(/Word document/)
  })
})
