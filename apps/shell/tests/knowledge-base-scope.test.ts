import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { FileIndexStore } from '../src/main/file-index/store'
import {
  KB_LIST_CAP,
  listIndexedKnowledgeFiles,
  normalizeKbScope,
  searchKnowledgeBase,
} from '../src/main/knowledge-base'

/**
 * The Ask AI corpus scope: 'indexed' (default) reaches every file the local
 * index covers, 'starred' keeps the original hand-picked set. Search drops the
 * JS post-filter for the indexed scope (the store's own ranked page is the
 * answer), the read gate probes index membership, and the prompt listing is a
 * bounded window of starred-first files.
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
