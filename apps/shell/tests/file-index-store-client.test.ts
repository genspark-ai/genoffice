import { mkdtempSync, rmSync, writeFileSync, unlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { StoreClient } from '../src/main/file-index/store-client'

/**
 * StoreClient (src/main/file-index/store-client.ts) against the mjs worker
 * doubles in this directory: the request/response roundtrip, error
 * propagation, the timeout that retires a wedged worker, and the worker that
 * exits before answering anything (the "store unavailable" state).
 */

const WORKER = join(__dirname, 'file-index-store-test-worker.mjs')
const EXIT_WORKER = join(__dirname, 'file-index-store-exit-worker.mjs')

let dir: string
/** `<dbPath>.hang`: while it exists the stub worker never answers */
const hangMarker = (): string => join(dir, 'index.db.hang')
const dbPath = (): string => join(dir, 'index.db')

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-store-client-'))
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('StoreClient', () => {
  it('round-trips a request and resolves its result', async () => {
    const client = new StoreClient(WORKER, dbPath(), 2_000)
    await expect(client.count()).resolves.toBe(7)
    // the worker hang marker never existed: a second call is served by the
    // same live worker
    await expect(client.count()).resolves.toBe(7)
    client.close()
  })

  it('rejects with the worker-reported error message', async () => {
    const client = new StoreClient(WORKER, dbPath(), 2_000)
    await expect(
      client.upsert({ path: '/a.md', mtimeMs: 1, sizeBytes: 1 }, 'text', 'ok'),
    ).rejects.toThrow('boom from the stub')
    client.close()
  })

  it('fails a wedged request on timeout, then serves the next call from a fresh worker', async () => {
    const client = new StoreClient(WORKER, dbPath(), 150)
    writeFileSync(hangMarker(), '1')
    await expect(client.count()).rejects.toThrow(/timed out/)
    // the wedged worker is retired; with the hang marker gone the next call
    // spawns a fresh one that answers
    unlinkSync(hangMarker())
    await expect(client.count()).resolves.toBe(7)
    client.close()
  })

  it('reports the store unavailable for good when the worker exits before answering', async () => {
    const client = new StoreClient(EXIT_WORKER, dbPath(), 2_000)
    await expect(client.count()).rejects.toThrow(/worker exited/)
    // no respawn loop against a DB that cannot even be opened
    await expect(client.count()).rejects.toThrow(/unavailable/)
    client.close()
  })
})
