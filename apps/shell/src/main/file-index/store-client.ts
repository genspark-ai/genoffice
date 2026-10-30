import { Worker } from 'node:worker_threads'
import type { IndexedFile, IndexStatus, SearchOptions, SearchResult } from './store'
import type { StoreRequest, StoreResponse } from './store-worker'
import type { IndexStore } from './index-store'
import type { ScannedFile } from './scan'

/** a request left unanswered this long means a wedged query or a dead worker:
 *  it fails, the worker is retired, and the next call spawns a fresh one (the
 *  WAL journal keeps the index file intact across the recycle) */
const REQUEST_TIMEOUT_MS = 30_000

/**
 * Main-process half of the store worker: an async IndexStore whose every call
 * is one request/response hop. The worker is spawned lazily on the first call
 * and recycled (not just retired) after a crash or a timeout — a store worker
 * that exits before answering anything marks this client dead for good, which
 * is the "store unavailable" state a failed in-process construction used to
 * produce.
 */
export class StoreClient implements IndexStore {
  private worker: Worker | null = null
  private nextId = 1
  private readonly waiting = new Map<
    number,
    { resolve: (value: unknown) => void; reject: (error: Error) => void }
  >()
  /** request ids posted to each worker: a retired worker's exit fails only its own */
  private readonly postedTo = new WeakMap<Worker, Set<number>>()
  /** a worker exited before answering anything: the DB cannot even be opened */
  private dead = false
  /** answered at least one request — a later exit is a crash, not a bad open */
  private answered = false

  constructor(
    private readonly workerPath: string,
    private readonly dbPath: string,
    private readonly requestTimeoutMs = REQUEST_TIMEOUT_MS,
  ) {}

  private ensureWorker(): Worker {
    if (this.dead) throw new Error('the file-index store is unavailable')
    if (this.worker) return this.worker
    const w = new Worker(this.workerPath, { workerData: { dbPath: this.dbPath } })
    w.on('message', (msg: StoreResponse) => {
      // any answer — result or error — proves the worker is alive (only a
      // worker that never answers was unable to open the DB)
      this.answered = true
      const pending = this.waiting.get(msg.id)
      if (!pending) return
      this.waiting.delete(msg.id)
      if (msg.ok) {
        pending.resolve(msg.result)
      } else {
        pending.reject(new Error(msg.error))
      }
    })
    const drop = (): void => {
      // in-flight requests posted to THIS worker fail with a named reason;
      // the next call respawns unless this worker died before ever answering
      // (an unopenable DB), and a successor's in-flight requests are safe —
      // its exit is not this one's to fail
      if (this.worker === w) this.worker = null
      if (!this.answered) this.dead = true
      const posted = this.postedTo.get(w) ?? new Set<number>()
      for (const id of posted) {
        const pending = this.waiting.get(id)
        if (!pending) continue
        this.waiting.delete(id)
        pending.reject(new Error('the file-index store worker exited'))
      }
      posted.clear()
    }
    w.on('error', drop)
    w.on('exit', drop)
    this.postedTo.set(w, new Set())
    this.worker = w
    return w
  }

  /** retire the current worker (timeout path); its exit handler fails the rest */
  private retireWorker(): void {
    const w = this.worker
    if (!w) return
    this.worker = null
    void w.terminate()
  }

  private call<T>(make: (id: number) => StoreRequest): Promise<T> {
    const id = this.nextId++
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.waiting.delete(id)
        this.retireWorker()
        reject(new Error('the file-index store request timed out'))
      }, this.requestTimeoutMs)
      this.waiting.set(id, {
        resolve: (value) => {
          clearTimeout(timer)
          resolve(value as T)
        },
        reject: (error) => {
          clearTimeout(timer)
          reject(error)
        },
      })
      try {
        const w = this.ensureWorker()
        this.postedTo.get(w)?.add(id)
        w.postMessage(make(id))
      } catch (e) {
        this.waiting.delete(id)
        clearTimeout(timer)
        reject(e instanceof Error ? e : new Error(String(e)))
      }
    })
  }

  search(input: string, opts?: SearchOptions): Promise<SearchResult> {
    return this.call<SearchResult>((id) => ({ id, type: 'search', input, opts }))
  }

  upsert(
    meta: { path: string; mtimeMs: number; sizeBytes: number },
    text: string | null,
    status: IndexStatus,
  ): Promise<void> {
    return this.call<void>((id) => ({ id, type: 'upsert', meta, text, status }))
  }

  remove(paths: readonly string[]): Promise<void> {
    return this.call<void>((id) => ({ id, type: 'remove', paths: [...paths] }))
  }

  diff(seen: readonly ScannedFile[], complete: boolean): Promise<ScannedFile[]> {
    return this.call<ScannedFile[]>((id) => ({ id, type: 'diff', seen: [...seen], complete }))
  }

  count(): Promise<number> {
    return this.call<number>((id) => ({ id, type: 'count' }))
  }

  has(path: string): Promise<boolean> {
    return this.call<boolean>((id) => ({ id, type: 'has', path }))
  }

  recentFiles(limit: number): Promise<IndexedFile[]> {
    return this.call<IndexedFile[]>((id) => ({ id, type: 'recentFiles', limit }))
  }

  excerptsFor(
    paths: readonly string[],
    q: string,
    chars: number,
  ): Promise<Array<{ path: string; name: string; excerpt: string }>> {
    return this.call<Array<{ path: string; name: string; excerpt: string }>>((id) => ({
      id,
      type: 'excerptsFor',
      paths: [...paths],
      q,
      chars,
    }))
  }

  /** will-quit hygiene: terminate the worker; in-flight calls fail via exit */
  close(): void {
    this.retireWorker()
  }
}
