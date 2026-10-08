import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { FileIndexer } from '../src/main/file-index/indexer'
import { FileIndexStore } from '../src/main/file-index/store'

const WORKER = join(__dirname, 'file-index-truncated-worker.mjs')

let indexer: FileIndexer | null
afterEach(() => {
  indexer?.stop()
  indexer = null
})

describe('FileIndexer with a budget-capped scan', () => {
  it('keeps indexed entries the capped walk never reached', async () => {
    const storeDir = mkdtempSync(join(tmpdir(), 'genoffice-indexer-db-'))
    const store = new FileIndexStore(join(storeDir, 'index.db'))
    try {
      const unreached = join(storeDir, 'deep', 'notes.md')
      store.upsert({ path: unreached, mtimeMs: 1, sizeBytes: 1 }, 'body', 'ok')
      indexer = new FileIndexer(store, WORKER, { roots: () => [storeDir], extraPaths: () => [] })
      await indexer.scan()
      expect(store.listAll().has(unreached)).toBe(true)
    } finally {
      store.close()
      rmSync(storeDir, { recursive: true, force: true })
    }
  })
})
