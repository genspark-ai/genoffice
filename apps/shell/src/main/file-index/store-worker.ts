/**
 * Worker-thread entry: the SQLite file index lives here so no search or write
 * blocks the main process — the one process serving every window's IPC. The
 * store opens from workerData at startup; a failed open kills the worker
 * before it answers anything, which the client reads as "the store is
 * unavailable" (the same outcome a failed in-process construction used to
 * have). Requests are answered one at a time in arrival order; a request that
 * throws answers with its error instead of taking the worker down.
 */
import { parentPort, workerData } from 'node:worker_threads'
import { diffAgainstKnown, FileIndexStore } from './store'
import type { IndexStatus, SearchOptions } from './store'
import type { ScannedFile } from './scan'

export type StoreRequest =
  | { id: number; type: 'search'; input: string; opts?: SearchOptions }
  | {
      id: number
      type: 'upsert'
      meta: { path: string; mtimeMs: number; sizeBytes: number }
      text: string | null
      status: IndexStatus
    }
  | { id: number; type: 'remove'; paths: string[] }
  | { id: number; type: 'diff'; seen: ScannedFile[]; complete: boolean }
  | { id: number; type: 'count' }
  | { id: number; type: 'has'; path: string }
  | { id: number; type: 'recentFiles'; limit: number }
  | { id: number; type: 'excerptsFor'; paths: string[]; q: string; chars: number }

export type StoreResponse =
  { id: number; ok: true; result: unknown } | { id: number; ok: false; error: string }

/**
 * One request, answered synchronously against the store it is handed —
 * exported so tests drive the exact dispatch the worker runs (the worker body
 * below is the same call wired to parentPort).
 */
export function answerStoreRequest(store: FileIndexStore, req: StoreRequest): unknown {
  switch (req.type) {
    case 'search':
      return store.search(req.input, req.opts)
    case 'upsert':
      store.upsert(req.meta, req.text, req.status)
      return undefined
    case 'remove':
      store.remove(req.paths)
      return undefined
    case 'diff':
      // listAll over the whole index runs here, not in the main process
      return diffAgainstKnown(store, req.seen, req.complete)
    case 'count':
      return store.count()
    case 'has':
      return store.has(req.path)
    case 'recentFiles':
      return store.recentFiles(req.limit)
    case 'excerptsFor':
      return store.excerptsFor(req.paths, req.q, req.chars)
  }
}

const dbPath = (workerData as { dbPath?: string } | undefined)?.dbPath
if (dbPath && parentPort) {
  const store = new FileIndexStore(dbPath)
  parentPort.on('message', (req: StoreRequest) => {
    try {
      const result = answerStoreRequest(store, req)
      parentPort!.postMessage({ id: req.id, ok: true, result } satisfies StoreResponse)
    } catch (e) {
      parentPort!.postMessage({
        id: req.id,
        ok: false,
        error: e instanceof Error ? e.message : String(e),
      } satisfies StoreResponse)
    }
  })
}
