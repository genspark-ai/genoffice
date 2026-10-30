import type { IndexedFile, IndexStatus, SearchOptions, SearchResult } from './store'
import { diffAgainstKnown, FileIndexStore } from './store'
import type { ScannedFile } from './scan'

/**
 * The file index as the main process sees it: FileIndexStore's surface with
 * every method async, because the store lives in a worker thread — a search or
 * write crossing this interface is a worker hop, never a synchronous SQLite
 * call on the process that serves every window's IPC. The worker body answers
 * each request; tests use the in-process adapter below to keep the same
 * contract without a thread.
 */
export interface IndexStore {
  search(input: string, opts?: SearchOptions): Promise<SearchResult>
  upsert(
    meta: { path: string; mtimeMs: number; sizeBytes: number },
    text: string | null,
    status: IndexStatus,
  ): Promise<void>
  remove(paths: readonly string[]): Promise<void>
  /**
   * Diff a completed walk against the index and answer the files that need
   * extraction; vanished entries are removed when the walk was complete. Runs
   * store-side so the whole-index listAll never materializes on the caller.
   */
  diff(seen: readonly ScannedFile[], complete: boolean): Promise<ScannedFile[]>
  count(): Promise<number>
  has(path: string): Promise<boolean>
  recentFiles(limit: number): Promise<IndexedFile[]>
  excerptsFor(
    paths: readonly string[],
    q: string,
    chars: number,
  ): Promise<Array<{ path: string; name: string; excerpt: string }>>
  /** release the store's resources (the worker's terminate / the db handle) */
  close(): void
}

/** the same store, answered in-process — the seam tests and any embedder that
 *  cannot spawn a worker use; behavior-identical to the worker path */
export function inProcessIndexStore(store: FileIndexStore): IndexStore {
  return {
    search: (input, opts) => Promise.resolve(store.search(input, opts)),
    upsert: (meta, text, status) => {
      store.upsert(meta, text, status)
      return Promise.resolve()
    },
    remove: (paths) => {
      store.remove(paths)
      return Promise.resolve()
    },
    diff: (seen, complete) => Promise.resolve(diffAgainstKnown(store, seen, complete)),
    count: () => Promise.resolve(store.count()),
    has: (path) => Promise.resolve(store.has(path)),
    recentFiles: (limit) => Promise.resolve(store.recentFiles(limit)),
    excerptsFor: (paths, q, chars) => Promise.resolve(store.excerptsFor(paths, q, chars)),
    close: () => store.close(),
  }
}
