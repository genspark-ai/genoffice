import { mkdirSync, mkdtempSync, rmSync, unlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { FileIndexer } from '../src/main/file-index/indexer'
import { FileIndexStore } from '../src/main/file-index/store'
import { inProcessIndexStore } from '../src/main/file-index/index-store'

/**
 * The query-freshness half of the indexer (src/main/file-index/indexer.ts):
 * settle() rides out an in-flight scan or extraction queue — bounded by its
 * cap — and indexNow() brings a first-sight file (a star, an open) into the
 * index without waiting for the next scan. The worker double hangs
 * extractions while a `hold` marker sits beside the target file, so a test
 * can hold the queue open and release it on cue.
 */
const WORKER = join(__dirname, 'file-index-settle-worker.mjs')

let dir: string
let storeDir: string
let store: FileIndexStore
let indexer: FileIndexer | null

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-settle-'))
  storeDir = mkdtempSync(join(tmpdir(), 'genoffice-settle-db-'))
  mkdirSync(join(dir, 'good'))
  writeFileSync(join(dir, 'good', 'notes.md'), 'searchable body text')
  store = new FileIndexStore(join(storeDir, 'index.db'))
})
afterEach(() => {
  indexer?.stop()
  indexer = null
  store.close()
  rmSync(dir, { recursive: true, force: true })
  rmSync(storeDir, { recursive: true, force: true })
})

const notesPath = () => join(dir, 'good', 'notes.md')
const holdPath = () => join(dir, 'good', 'hold.md')
const holdMarker = () => join(dir, 'good', 'hold')

const makeIndexer = (workerIdleMs = 60_000) =>
  new FileIndexer(
    inProcessIndexStore(store),
    WORKER,
    { roots: () => [dir], extraPaths: () => [] },
    30_000,
    256,
    workerIdleMs, // long by default: retire-on-idle would race the settle waits
  )

describe('FileIndexer.settle', () => {
  it('answers immediately when the pipeline is quiet', async () => {
    indexer = makeIndexer()
    const started = Date.now()
    await indexer.settle(5_000)
    expect(Date.now() - started).toBeLessThan(500)
  })

  it('waits out a draining extraction queue, then the file is in the index', async () => {
    indexer = makeIndexer()
    writeFileSync(holdMarker(), '1')
    writeFileSync(holdPath(), 'held back by the marker')
    await indexer.indexNow([holdPath()])
    // the queue is open and the worker is holding the extraction
    const settled = indexer.settle(10_000)
    let resolved = false
    void settled.then(() => {
      resolved = true
    })
    await vi.waitFor(() => expect(indexer!.workerIsRunning()).toBe(true))
    expect(resolved).toBe(false)
    // releasing the marker lets the extraction land, and only then settles
    unlinkSync(holdMarker())
    await settled
    expect(store.has(holdPath())).toBe(true)
  })

  it('answers at its cap when the pipeline never goes quiet', async () => {
    indexer = makeIndexer()
    writeFileSync(holdMarker(), '1')
    writeFileSync(holdPath(), 'held back by the marker')
    await indexer.indexNow([holdPath()])
    const started = Date.now()
    await indexer.settle(150)
    const waited = Date.now() - started
    expect(waited).toBeGreaterThanOrEqual(140)
    expect(waited).toBeLessThan(2_000)
    // the wedged extraction did not land; the cap answered instead
    expect(store.has(holdPath())).toBe(false)
    unlinkSync(holdMarker())
  })

  it('rides out the refresh debounce and the scan behind it (the just-saved file)', async () => {
    indexer = makeIndexer()
    // the fresh save the watcher would report: at the scan root, not indexed yet
    const saved = join(dir, 'saved.md')
    writeFileSync(saved, 'just saved')
    const started = Date.now()
    indexer.refresh()
    // without settle a query right now answers from the pre-save index
    expect(store.has(saved)).toBe(false)
    await indexer.settle(10_000)
    // the debounce (1.5s) plus the scan passed before settle answered
    expect(Date.now() - started).toBeGreaterThanOrEqual(1_400)
    expect(store.has(saved)).toBe(true)
  })

  it('stop() releases a waiting settle instead of leaving it to the cap', async () => {
    indexer = makeIndexer()
    writeFileSync(holdMarker(), '1')
    writeFileSync(holdPath(), 'held back by the marker')
    await indexer.indexNow([holdPath()])
    const settled = indexer.settle(10_000)
    await vi.waitFor(() => expect(indexer!.workerIsRunning()).toBe(true))
    indexer!.stop()
    const started = Date.now()
    await settled
    expect(Date.now() - started).toBeLessThan(500)
  })
})

describe('FileIndexer.indexNow', () => {
  it('extracts a first-sight file the index does not carry yet', async () => {
    indexer = makeIndexer()
    await indexer.indexNow([notesPath()])
    await indexer.settle(5_000)
    expect(store.has(notesPath())).toBe(true)
  })

  it('skips a file the index already carries (the scan owns changes)', async () => {
    // a short idle recycle so the worker demonstrably retires after the first
    // extraction; the second indexNow must not wake one up
    indexer = makeIndexer(150)
    await indexer.indexNow([notesPath()])
    await indexer.settle(5_000)
    await vi.waitFor(() => expect(indexer!.workerIsRunning()).toBe(false), {
      timeout: 2_000,
      interval: 25,
    })
    // a second indexNow of the same path must not spawn a worker or queue work
    await indexer.indexNow([notesPath()])
    await new Promise((resolve) => setTimeout(resolve, 150))
    expect(indexer.workerIsRunning()).toBe(false)
    await expect(indexer.progress()).resolves.toMatchObject({ pending: 0 })
  })

  it('ignores unsupported extensions and paths that do not exist', async () => {
    indexer = makeIndexer()
    writeFileSync(join(dir, 'setup.exe'), 'binary noise')
    await indexer.indexNow([join(dir, 'setup.exe'), join(dir, 'good', 'ghost.md')])
    await new Promise((resolve) => setTimeout(resolve, 150))
    expect(indexer.workerIsRunning()).toBe(false)
    await expect(indexer.progress()).resolves.toMatchObject({ pending: 0 })
  })
})
