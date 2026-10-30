import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { FileIndexStore } from '../src/main/file-index/store'
import { answerStoreRequest, type StoreRequest } from '../src/main/file-index/store-worker'

/**
 * The store worker's dispatch (answerStoreRequest): every request type runs
 * against a real FileIndexStore, exactly as the worker thread runs it — the
 * protocol shapes are what StoreClient on the main side decodes.
 */

let dir: string
let store: FileIndexStore

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-store-worker-'))
  store = new FileIndexStore(join(dir, 'index.db'))
})

afterEach(() => {
  store.close()
  rmSync(dir, { recursive: true, force: true })
})

const ask = (req: Omit<StoreRequest, 'id'>): unknown =>
  answerStoreRequest(store, { id: 1, ...req } as StoreRequest)

describe('answerStoreRequest', () => {
  it('answers search, upsert, count, has and recentFiles against the store', () => {
    expect(ask({ type: 'count' })).toBe(0)
    ask({
      type: 'upsert',
      meta: { path: '/a.md', mtimeMs: 1, sizeBytes: 2 },
      text: 'hello world',
      status: 'ok',
    })
    expect(ask({ type: 'count' })).toBe(1)
    expect(ask({ type: 'has', path: '/a.md' })).toBe(true)
    expect(ask({ type: 'has', path: '/b.md' })).toBe(false)
    const search = ask({ type: 'search', input: 'hello' }) as { hits: Array<{ path: string }> }
    expect(search.hits.map((h) => h.path)).toEqual(['/a.md'])
    const recent = ask({ type: 'recentFiles', limit: 10 }) as Array<{ path: string }>
    expect(recent.map((r) => r.path)).toEqual(['/a.md'])
  })

  it('answers excerptsFor with name and excerpt text', () => {
    ask({
      type: 'upsert',
      meta: { path: '/a.md', mtimeMs: 1, sizeBytes: 2 },
      text: 'the needle is here',
      status: 'ok',
    })
    const excerpts = ask({
      type: 'excerptsFor',
      paths: ['/a.md'],
      q: 'needle',
      chars: 40,
    }) as Array<{
      path: string
      name: string
      excerpt: string
    }>
    expect(excerpts).toHaveLength(1)
    expect(excerpts[0]!.name).toBe('a.md')
    expect(excerpts[0]!.excerpt).toContain('needle')
  })

  it('answers diff with the changed files and removes the vanished', () => {
    ask({
      type: 'upsert',
      meta: { path: '/old.md', mtimeMs: 1, sizeBytes: 1 },
      text: 'x',
      status: 'ok',
    })
    ask({
      type: 'upsert',
      meta: { path: '/gone.md', mtimeMs: 1, sizeBytes: 1 },
      text: 'x',
      status: 'ok',
    })
    const changed = ask({
      type: 'diff',
      seen: [
        // changed on disk → answered for extraction (and kept)
        { path: '/old.md', mtimeMs: 2, sizeBytes: 1 },
        // never indexed → answered
        { path: '/new.md', mtimeMs: 5, sizeBytes: 5 },
        // /gone.md is absent from a complete walk → removed
      ],
      complete: true,
    }) as Array<{ path: string }>
    expect(changed.map((f) => f.path)).toEqual(['/old.md', '/new.md'])
    expect(ask({ type: 'has', path: '/gone.md' })).toBe(false)
    expect(ask({ type: 'has', path: '/old.md' })).toBe(true)
    // an incomplete walk keeps entries the walk never reached
    ask({
      type: 'upsert',
      meta: { path: '/kept.md', mtimeMs: 1, sizeBytes: 1 },
      text: 'x',
      status: 'ok',
    })
    ask({ type: 'diff', seen: [], complete: false })
    expect(ask({ type: 'has', path: '/kept.md' })).toBe(true)
  })

  it('answers remove by path', () => {
    ask({
      type: 'upsert',
      meta: { path: '/a.md', mtimeMs: 1, sizeBytes: 1 },
      text: 'x',
      status: 'ok',
    })
    ask({ type: 'remove', paths: ['/a.md'] })
    expect(ask({ type: 'has', path: '/a.md' })).toBe(false)
  })
})
