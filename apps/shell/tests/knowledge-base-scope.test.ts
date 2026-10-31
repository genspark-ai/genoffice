import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { FileIndexStore } from '../src/main/file-index/store'
import {
  KB_LIST_CAP,
  KB_ONLY_PATHS_MAX,
  kbReadAllows,
  listIndexedKnowledgeFiles,
  normalizeKbScope,
  notIndexedPaths,
  parseKbOnlyPaths,
  searchKnowledgeBase,
} from '../src/main/knowledge-base'

/**
 * The Ask AI corpus scope: 'indexed' (default) reaches every file the local
 * index covers, 'starred' keeps the original hand-picked set. Search drops the
 * JS post-filter for the indexed scope (the store's own ranked page is the
 * answer), the read gate probes index membership, and the prompt listing is a
 * bounded window of starred-first files.
 *
 * The round scope (an explicit file selection as one conversation's corpus)
 * rides on top: onlyPaths payloads are validated (strings, absolute, length,
 * count), the read gate narrows to the selection and admits a selected
 * existing office file only at a whitelisted location (a folder root, a
 * desktop alias, a recent or a star), and the not-indexed report says which
 * selected paths the index does not cover.
 */

const HIT = (path: string) => ({
  path,
  name: path.split('/').pop()!,
  ext: 'docx',
  snippetText: '…hit…',
})

describe('normalizeKbScope', () => {
  it('defaults to indexed, keeps starred, rejects junk', () => {
    expect(normalizeKbScope(undefined)).toBe('indexed')
    expect(normalizeKbScope('starred')).toBe('starred')
    expect(normalizeKbScope('everything')).toBe('indexed')
  })
})

describe('searchKnowledgeBase', () => {
  const poolOf = (paths: string[]) => paths.map(HIT)

  it('indexed scope (null) takes the store page directly — no 200-pool fetch', async () => {
    const requested: number[] = []
    const search = (q: string, limit: number) => {
      requested.push(limit)
      return poolOf(['/a/1.docx', '/b/2.docx'])
    }
    const hits = await searchKnowledgeBase(async (q, limit) => search(q, limit), 'report', null, 8)
    expect(requested).toEqual([8])
    expect(hits.map((h) => h.path)).toEqual(['/a/1.docx', '/b/2.docx'])
  })

  it('starred scope keeps the pool fetch and the membership filter', async () => {
    const requested: number[] = []
    const search = (q: string, limit: number) => {
      requested.push(limit)
      return poolOf(['/a/starred.docx', '/b/other.docx'])
    }
    const hits = await searchKnowledgeBase(
      async (q, limit) => search(q, limit),
      'report',
      ['/a/starred.docx'],
      8,
    )
    expect(requested).toEqual([200])
    expect(hits.map((h) => h.path)).toEqual(['/a/starred.docx'])
  })

  it('trims and de-duplicates in both scopes', async () => {
    const search = () => poolOf(['/a/1.docx', '/a/1.docx', '/b/2.docx'])
    expect(await searchKnowledgeBase(async () => search(), '  x  ', null, 5)).toHaveLength(2)
    expect(await searchKnowledgeBase(async () => search(), '   ', null, 5)).toEqual([])
  })
})

describe('parseKbOnlyPaths — the round-scope payload', () => {
  it('absent and empty mean no round scope (null)', () => {
    expect(parseKbOnlyPaths(undefined)).toBeNull()
    expect(parseKbOnlyPaths(null)).toBeNull()
    expect(parseKbOnlyPaths([])).toBeNull()
  })

  it('keeps a validated selection, deduplicated in order', () => {
    expect(parseKbOnlyPaths(['/a/1.docx', '/a/1.docx', '/b/2.docx'])).toEqual([
      '/a/1.docx',
      '/b/2.docx',
    ])
  })

  it('rejects wrong shapes, relative paths, over-long paths and over-many paths', () => {
    expect(() => parseKbOnlyPaths('/a/1.docx')).toThrow(/array of absolute paths/)
    expect(() => parseKbOnlyPaths({ 0: '/a/1.docx' })).toThrow(/array of absolute paths/)
    expect(() => parseKbOnlyPaths(['/a/1.docx', 42])).toThrow(/must be strings/)
    expect(() => parseKbOnlyPaths(['relative/1.docx'])).toThrow(/absolute/)
    expect(() => parseKbOnlyPaths([''])).toThrow(/absolute/)
    expect(() => parseKbOnlyPaths([`/${'a'.repeat(4097)}.docx`])).toThrow(/too long/)
    expect(() =>
      parseKbOnlyPaths(Array.from({ length: KB_ONLY_PATHS_MAX + 1 }, (_, i) => `/idx/${i}.docx`)),
    ).toThrow(/at most 50/)
    // node:path's isAbsolute is platform-specific: a Windows drive path is
    // absolute only where it is a real path
    const winPath = 'C:\\Users\\me\\a.docx'
    if (process.platform === 'win32') {
      expect(parseKbOnlyPaths([winPath])).toEqual([winPath])
    } else {
      expect(() => parseKbOnlyPaths([winPath])).toThrow(/absolute/)
    }
  })
})

describe('kbReadAllows — the read gate with a round scope', () => {
  let dir: string
  const statFile = (path: string) => {
    try {
      return statSync(path)
    } catch {
      return null
    }
  }

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-gate-'))
  })
  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  it('admits what the corpus already admits, scope or not', () => {
    expect(
      kbReadAllows('/idx/a.docx', {
        inCorpus: true,
        scopePaths: null,
        locationAllowed: false,
        statFile,
      }),
    ).toBe(true)
    expect(
      kbReadAllows('/idx/a.docx', {
        inCorpus: false,
        scopePaths: null,
        locationAllowed: true,
        statFile,
      }),
    ).toBe(false)
  })

  it('a scope narrows: a corpus path outside the selection is refused, an in-scope one reads', () => {
    const indexed = join(dir, 'indexed.docx')
    writeFileSync(indexed, 'x')
    expect(
      kbReadAllows(indexed, {
        inCorpus: true,
        scopePaths: [join(dir, 'picked.docx')],
        locationAllowed: false,
        statFile,
      }),
    ).toBe(false)
    expect(
      kbReadAllows(indexed, {
        inCorpus: true,
        scopePaths: [indexed],
        locationAllowed: false,
        statFile,
      }),
    ).toBe(true)
  })

  it('a selected file the index missed is readable when its location is whitelisted and it is an office type', () => {
    const picked = join(dir, 'picked.docx')
    writeFileSync(picked, 'x')
    expect(
      kbReadAllows(picked, {
        inCorpus: false,
        scopePaths: [picked],
        locationAllowed: true,
        statFile,
      }),
    ).toBe(true)
  })

  it('refuses a selected path whose location is not whitelisted — off every root, not an alias, not tracked', () => {
    const stray = join(dir, 'stray.docx')
    writeFileSync(stray, 'x')
    expect(
      kbReadAllows(stray, {
        inCorpus: false,
        scopePaths: [stray],
        locationAllowed: false,
        statFile,
      }),
    ).toBe(false)
  })

  it('refuses a whitelisted selected path that is gone, a directory, or not a readable office type', () => {
    const gone = join(dir, 'gone.docx')
    const archive = join(dir, 'packed.zip')
    writeFileSync(archive, 'not office')
    const folder = join(dir, 'sub.docx')
    mkdirSync(folder)
    for (const path of [gone, archive, folder]) {
      expect(
        kbReadAllows(path, {
          inCorpus: false,
          scopePaths: [path],
          locationAllowed: true,
          statFile,
        }),
      ).toBe(false)
    }
  })

  it('refuses a path outside the selection even when a scope is attached', () => {
    const picked = join(dir, 'picked.docx')
    writeFileSync(picked, 'x')
    expect(
      kbReadAllows(join(dir, 'other.docx'), {
        inCorpus: false,
        scopePaths: [picked],
        locationAllowed: true,
        statFile,
      }),
    ).toBe(false)
  })
})

describe('notIndexedPaths — the scope paths the index does not cover', () => {
  it('answers membership per path, in order', async () => {
    const store = { has: async (path: string) => path.startsWith('/idx/') }
    expect(
      await notIndexedPaths(store, ['/idx/a.docx', '/elsewhere/b.docx', '/idx/c.docx']),
    ).toEqual(['/elsewhere/b.docx'])
  })

  it('a failing membership probe reads as not indexed, no throw', async () => {
    const store = {
      has: async (path: string) => {
        if (path === '/boom') throw new Error('store gone')
        return true
      },
    }
    expect(await notIndexedPaths(store, ['/fine.docx', '/boom'])).toEqual(['/boom'])
  })

  it('no store: the whole selection reports not indexed', async () => {
    expect(await notIndexedPaths(null, ['/a.docx', '/b.docx'])).toEqual(['/a.docx', '/b.docx'])
  })
})

describe('listIndexedKnowledgeFiles', () => {
  let dir: string
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-scope-'))
  })
  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  const indexed = (paths: string[]) =>
    paths.map((path, i) => ({ path, mtimeMs: i, sizeBytes: 10 + i, status: 'ok' as const }))

  it('lists starred files first, then recent indexed ones, deduplicated and capped', () => {
    const starred = join(dir, 'starred.docx')
    writeFileSync(starred, 'x')
    const list = listIndexedKnowledgeFiles(
      [starred, join(dir, 'gone.docx')],
      indexed([starred, '/idx/a.docx', '/idx/b.docx']),
    )
    // the starred file exists on disk and leads; the vanished one drops out;
    // the indexed copy of the starred path is not duplicated
    expect(list.map((f) => f.path)).toEqual([starred, '/idx/a.docx', '/idx/b.docx'])
    expect(list[0]!.folder).toBe(basename(dir))
  })

  it('honors the cap', () => {
    const many = indexed(Array.from({ length: 300 }, (_, i) => `/idx/file-${i}.docx`))
    expect(listIndexedKnowledgeFiles([], many)).toHaveLength(KB_LIST_CAP)
  })
})

describe('FileIndexStore.has / recentFiles (the indexed-scope gates)', () => {
  let dir: string
  let store: FileIndexStore

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'genoffice-kb-index-'))
    store = new FileIndexStore(join(dir, 'index.db'))
  })
  afterEach(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })

  it('has() answers membership exactly', () => {
    store.upsert({ path: '/docs/a.docx', mtimeMs: 1, sizeBytes: 10 }, 'hello', 'ok')
    expect(store.has('/docs/a.docx')).toBe(true)
    expect(store.has('/docs/other.docx')).toBe(false)
  })

  it('recentFiles() returns newest first, honoring the limit', () => {
    store.upsert({ path: '/docs/old.docx', mtimeMs: 100, sizeBytes: 10 }, null, 'name-only')
    store.upsert({ path: '/docs/new.docx', mtimeMs: 200, sizeBytes: 20 }, 'text', 'ok')
    store.upsert({ path: '/docs/mid.docx', mtimeMs: 150, sizeBytes: 15 }, 'text', 'ok')
    const all = store.recentFiles(10)
    expect(all.map((f) => f.path)).toEqual(['/docs/new.docx', '/docs/mid.docx', '/docs/old.docx'])
    expect(all[0]!.sizeBytes).toBe(20)
    expect(store.recentFiles(2)).toHaveLength(2)
  })
})
