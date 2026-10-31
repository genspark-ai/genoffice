import { mkdtempSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * The knowledge-base channels behind the Ask AI panel (kb:search / kb:read in
 * home-ipc) against a mocked Electron and a fake index store: the round scope
 * — an explicit file selection as one conversation's corpus — intersects
 * search with the selection, reports the selected paths the index does not
 * cover as `notIndexed`, and narrows the read gate the same way: a selected
 * existing office file is readable only at a whitelisted location (the alias
 * list, the recents, the stars, or inside a folder root — the host and
 * registry stubs answer those lists per test), and a corpus file outside the
 * selection is refused. Scope payloads are validated the same way in both
 * channels: strings, absolute, per-path length, count. Unscoped calls keep
 * answering the bare hit array a renderer that predates onlyPaths reads.
 */

const handles = new Map<string, (...args: never[]) => unknown>()

vi.mock('electron', () => ({
  app: { getPath: () => '/tmp/genoffice-home-ipc-kb' },
  BrowserWindow: class {},
  dialog: {},
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      handles.set(channel, handler),
  },
  shell: {},
}))

// app-registry re-exports the six app main modules; the kb handlers only need
// the starred and recent lists, so the registry is stubbed rather than loaded.
// The descriptor table stays (app-routing builds its extension regexes from it at
// import time) with just the fields that load-time code reads. The lists are
// hoisted mutable state so a test can place a path in the round-scope
// location whitelist (recents) or leave it out.
const APPS = vi.hoisted(() =>
  ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html'].map((id) => ({
    id,
    family: id,
    extensions: [id === 'markdown' ? 'md' : id],
    untitledTitle: id,
    untitledTitleKey: `untitled-${id}`,
  })),
)

/** the recents list the registry stub answers with (per test) */
const recentsState = vi.hoisted(() => ({ paths: [] as string[] }))
/** the desktop alias list the host stub answers with (per test) */
const aliasState = vi.hoisted(() => ({ paths: [] as string[] }))

vi.mock('../src/main/app-registry', () => ({
  documentApps: APPS,
  appForPath: (path: string) =>
    APPS.find((app) => path.endsWith(`.${app.extensions[0]}`))?.id ?? null,
  appForKind: (kind: string) => APPS.find((app) => app.id === kind) ?? null,
  appIsDirty: () => false,
  resetSheetsShuttingDown: () => {},
  defaultSaveDir: () => '/tmp/genoffice-home-ipc-kb/save',
  projectFileRenamed: () => {},
  readRecentFiles: () => recentsState.paths,
  readStarredFiles: () => [] as string[],
  readStarredGroupMap: () => new Map<string, string>(),
  readStarredGroups: () => [] as string[],
  recordRecentFile: () => {},
  removeRecentFiles: () => {},
  removeStarredFiles: () => {},
  setStarredGroup: () => {},
  toggleStarredFile: () => {},
  uniquePathIn: (dir: string, name: string) => join(dir, name),
}))

import { registerHomeFileIpc, type HomeFileIpcHost } from '../src/main/home-ipc'

/** the fake store: a fixed search table plus an index-membership set */
const INDEXED = new Set<string>(['/idx/contract.docx', '/idx/other.pdf'])

const searchImpl = vi.fn(async (q: string, opts?: { limit?: number }) => ({
  total: 3,
  hits: [...INDEXED]
    .filter((path) => path.includes(q))
    .slice(0, opts?.limit ?? 50)
    .map((path) => ({
      path,
      name: path.split('/').pop()!,
      ext: path.split('.').pop()!,
      folder: 'idx',
      snippet: [{ text: '…hit…' }],
      mtimeMs: 1,
      sizeBytes: 10,
      starred: false,
    })),
}))

const store = {
  search: searchImpl,
  has: async (path: string) => INDEXED.has(path),
  recentFiles: async () => [] as Array<{ path: string; mtimeMs: number; sizeBytes: number }>,
}

const host = {
  translate: (key: string) => key,
  appSettingsPath: () => settingsPath,
  getShellWindow: () => null,
  getTabManager: () => null,
  ensureFileIndexer: () => null,
  indexStore: () => store as never,
  refreshIndexer: () => {},
  indexNow: async () => {},
  readFileSearchSettings: () => ({ rerank: false }) as never,
  ensureSearchReranker: () => null,
  openDocumentPath: () => false,
  rememberPendingDir: () => {},
  newDocTab: () => {},
  newSheetTab: () => {},
  newSlideTab: () => {},
  newMarkdownTab: () => {},
  newHtmlTab: () => {},
  newPdfTab: () => {},
  afterFileMoved: () => {},
  afterFolderMoved: () => {},
  trackedFilesUnder: () => [] as string[],
  fileTargetSources: () => ({}) as never,
  statMaybeFile: (path: string) => {
    try {
      return statSync(path)
    } catch {
      return null
    }
  },
  extraFolderRoots: () => [] as string[],
  insideAnyRoot: () => false,
  isAnyRoot: () => false,
  desktopAliases: () => aliasState.paths,
  ensureFolderWatchers: () => {},
  droppedFilesDeps: () => ({}) as never,
  cliRunner: () => ({}) as never,
  fetchGithubStars: async () => null,
  openDialogExtensions: [] as readonly string[],
} as HomeFileIpcHost

const invoke = (channel: string, ...args: unknown[]) =>
  handles.get(channel)!({ sender: { id: 1 } } as never, ...args)

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-ipc-'))
  settingsPath = join(dir, 'app-settings.json')
  handles.clear()
  searchImpl.mockClear()
  recentsState.paths = []
  aliasState.paths = []
  registerHomeFileIpc(host)
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('kb:search — the round scope', () => {
  it('answers a bare hit array when no onlyPaths rides the query', async () => {
    const answer = await invoke('kb:search', { q: 'contract', limit: 8 })
    expect(Array.isArray(answer)).toBe(true)
    expect((answer as { path: string }[]).map((h) => h.path)).toEqual(['/idx/contract.docx'])
  })

  it('intersects the index with the selection and reports the not-indexed paths', async () => {
    const answer = (await invoke('kb:search', {
      q: 'idx',
      limit: 8,
      onlyPaths: ['/idx/contract.docx', '/idx/unindexed.docx', '/elsewhere/notes.md'],
    })) as { hits: { path: string }[]; notIndexed: string[] }
    expect(answer.hits.map((h) => h.path)).toEqual(['/idx/contract.docx'])
    // the second path exists in no index; the third is another index's
    expect(answer.notIndexed).toEqual(['/idx/unindexed.docx', '/elsewhere/notes.md'])
  })

  it('the round scope overrides a starred kbScope for that call', async () => {
    writeFileSync(settingsPath, JSON.stringify({ kbScope: 'starred' }))
    const answer = (await invoke('kb:search', {
      q: 'contract',
      limit: 8,
      onlyPaths: ['/idx/contract.docx'],
    })) as { hits: { path: string }[]; notIndexed: string[] }
    // unstarred and outside the starred corpus — in the selection, so a hit
    expect(answer.hits.map((h) => h.path)).toEqual(['/idx/contract.docx'])
  })

  it('no store answers an empty scope with every path not indexed', async () => {
    const hostNoStore = { ...host, indexStore: () => null }
    handles.clear()
    registerHomeFileIpc(hostNoStore)
    const answer = (await invoke('kb:search', {
      q: 'anything',
      limit: 8,
      onlyPaths: ['/idx/contract.docx'],
    })) as { hits: unknown[]; notIndexed: string[] }
    expect(answer.hits).toEqual([])
    expect(answer.notIndexed).toEqual(['/idx/contract.docx'])
  })

  it('a blank query with a scope keeps the scoped answer shape', async () => {
    const answer = (await invoke('kb:search', {
      q: '   ',
      limit: 8,
      onlyPaths: ['/idx/contract.docx', '/idx/unindexed.docx'],
    })) as { hits: unknown[]; notIndexed: string[] }
    expect(answer.hits).toEqual([])
    expect(answer.notIndexed).toEqual(['/idx/unindexed.docx'])
  })

  it('rejects malformed scope payloads before anything runs', async () => {
    await expect(invoke('kb:search', { q: 'x', onlyPaths: 'not-an-array' })).rejects.toThrow(
      /array of absolute paths/,
    )
    await expect(invoke('kb:search', { q: 'x', onlyPaths: ['/idx/a.docx', 42] })).rejects.toThrow(
      /must be strings/,
    )
    await expect(invoke('kb:search', { q: 'x', onlyPaths: ['relative/a.docx'] })).rejects.toThrow(
      /absolute/,
    )
    await expect(
      invoke('kb:search', { q: 'x', onlyPaths: [`/${'a'.repeat(4097)}.docx`] }),
    ).rejects.toThrow(/too long/)
    await expect(
      invoke('kb:search', {
        q: 'x',
        onlyPaths: Array.from({ length: 51 }, (_, i) => `/idx/${i}.docx`),
      }),
    ).rejects.toThrow(/at most 50/)
    // nothing above reached the store
    expect(searchImpl).not.toHaveBeenCalled()
  })
})

describe('kb:read — the gate', () => {
  it('reads an indexed file, refuses one off-index without a scope', async () => {
    const indexed = join(dir, 'indexed.md')
    writeFileSync(indexed, 'in the index')
    INDEXED.add(indexed)
    const page = (await invoke('kb:read', { path: indexed, offset: 0 })) as { text: string }
    expect(page.text).toBe('in the index')
    const offIndex = join(dir, 'off-index.md')
    writeFileSync(offIndex, 'never indexed')
    await expect(invoke('kb:read', { path: offIndex, offset: 0 })).rejects.toThrow(
      /not in the knowledge base/,
    )
  })

  it('a scope narrows: a corpus path outside the selection is refused', async () => {
    const picked = join(dir, 'picked.md')
    const other = join(dir, 'other.md')
    writeFileSync(picked, 'selected corpus file')
    writeFileSync(other, 'in the corpus, not selected')
    aliasState.paths = [picked]
    INDEXED.add(other)
    const page = (await invoke('kb:read', {
      path: picked,
      offset: 0,
      onlyPaths: [picked],
    })) as { text: string }
    expect(page.text).toBe('selected corpus file')
    // the chip means the selection is the corpus: index membership alone
    // (union) must not leak a path the request did not select
    await expect(
      invoke('kb:read', { path: other, offset: 0, onlyPaths: [picked] }),
    ).rejects.toThrow(/not in the knowledge base/)
  })

  it('a selected file the index missed stays readable when it sits on the alias list', async () => {
    const picked = join(dir, 'picked.md')
    writeFileSync(picked, 'selected corpus file')
    aliasState.paths = [picked]
    const page = (await invoke('kb:read', {
      path: picked,
      offset: 0,
      onlyPaths: [picked],
    })) as { text: string; offset: number }
    expect(page.text).toBe('selected corpus file')
    expect(page.offset).toBe(0)
  })

  it('a selected file the index missed stays readable when it sits among the recents', async () => {
    const recent = join(dir, 'recent.md')
    writeFileSync(recent, 'a recent, unindexed')
    recentsState.paths = [recent]
    const page = (await invoke('kb:read', {
      path: recent,
      offset: 0,
      onlyPaths: [recent],
    })) as { text: string }
    expect(page.text).toBe('a recent, unindexed')
  })

  it('a selected path outside every root and off every list is refused, even in onlyPaths', async () => {
    // the host answers no roots; the registry stub answers no recents or
    // stars, and the alias list is empty — a real office file in a scratch
    // directory is exactly the arbitrary-file read the gate must refuse
    const stray = join(dir, 'stray.md')
    writeFileSync(stray, 'nowhere the desktop could have selected me')
    await expect(invoke('kb:read', { path: stray, offset: 0, onlyPaths: [stray] })).rejects.toThrow(
      /not in the knowledge base/,
    )
  })

  it('a selected path is refused when the file is gone or not a readable office type', async () => {
    const gone = join(dir, 'gone.docx')
    await expect(
      invoke('kb:read', {
        path: gone,
        offset: 0,
        onlyPaths: [gone],
      }),
    ).rejects.toThrow(/not in the knowledge base/)
    const archive = join(dir, 'packed.zip')
    writeFileSync(archive, 'not office')
    aliasState.paths = [archive]
    await expect(
      invoke('kb:read', { path: archive, offset: 0, onlyPaths: [archive] }),
    ).rejects.toThrow(/not in the knowledge base/)
  })

  it('a file outside the selection is refused even when the request carries a scope', async () => {
    const picked = join(dir, 'picked.md')
    writeFileSync(picked, 'selected corpus file')
    aliasState.paths = [picked]
    const other = join(dir, 'other.md')
    writeFileSync(other, 'not selected')
    await expect(
      invoke('kb:read', { path: other, offset: 0, onlyPaths: [picked] }),
    ).rejects.toThrow(/not in the knowledge base/)
  })

  it('validates the read-side scope payload the same way', async () => {
    await expect(
      invoke('kb:read', { path: '/idx/contract.docx', offset: 0, onlyPaths: ['relative.md'] }),
    ).rejects.toThrow(/absolute/)
    await expect(
      invoke('kb:read', {
        path: '/idx/contract.docx',
        offset: 0,
        onlyPaths: Array.from({ length: 51 }, (_, i) => `/idx/${i}.md`),
      }),
    ).rejects.toThrow(/at most 50/)
  })
})
