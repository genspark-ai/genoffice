import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { FileIndexer } from '../src/main/file-index/indexer'
import { FileIndexStore } from '../src/main/file-index/store'

// the worker double never answers the extraction for poison.pdf, so these tests
// only pass once a wedged request fails on timeout and the queue moves on to
// the files behind it
const WORKER = join(__dirname, 'file-index-test-worker.mjs')

let dir: string
let storeDir: string
let store: FileIndexStore
let indexer: FileIndexer | null

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-indexer-'))
  storeDir = mkdtempSync(join(tmpdir(), 'genoffice-indexer-db-'))
  mkdirSync(join(dir, 'hang'))
  writeFileSync(join(dir, 'hang', 'poison.pdf'), 'unparseable')
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

const indexerWithShortTimeout = () =>
  new FileIndexer(store, WORKER, { roots: () => [dir], extraPaths: () => [] }, 400)

const notesPath = () => join(dir, 'good', 'notes.md')
const poisonPath = () => join(dir, 'hang', 'poison.pdf')

describe('FileIndexer wedged-worker recovery', () => {
  it('fails a hung extraction on timeout and keeps indexing the rest of the queue', async () => {
    indexer = indexerWithShortTimeout()
    const started = Date.now()
    await indexer.scan()
    await vi.waitFor(
      () => {
        expect(store.listAll().get(poisonPath())?.status).toBe('error')
      },
      { timeout: 5_000, interval: 25 },
    )
    const settleMs = Date.now() - started
    expect(store.listAll().get(notesPath())?.status).toBe('ok')
    expect(indexer.progress()).toMatchObject({ pending: 0, scanning: false })
    // the queue waited out the injected 400ms wedge, then finished
    expect(settleMs).toBeGreaterThanOrEqual(400)
    expect(settleMs).toBeLessThan(5_000)
    console.log('settle ms:', settleMs)
  })

  it('still runs scans and extractions after a timed-out request', async () => {
    indexer = indexerWithShortTimeout()
    await indexer.scan()
    await vi.waitFor(
      () => {
        expect(indexer!.progress().pending).toBe(0)
      },
      { timeout: 5_000, interval: 25 },
    )
    expect(store.listAll().get(notesPath())?.status).toBe('ok')
    // error rows are retried on every scan; the retry must not wedge either
    await indexer.scan()
    await vi.waitFor(
      () => {
        expect(indexer!.progress()).toMatchObject({ pending: 0, scanning: false })
      },
      { timeout: 5_000, interval: 25 },
    )
    expect(store.listAll().get(notesPath())?.status).toBe('ok')
  })
})
