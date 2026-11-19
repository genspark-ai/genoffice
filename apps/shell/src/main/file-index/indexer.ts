import { Worker } from 'node:worker_threads'
import type { Extracted } from './extract'
import type { WorkerRequest, WorkerResponse } from './extract-worker'
import { isSupportedTreeFile } from '../folder-tree'
import { statOrNull, type ScannedFile } from './scan'
import type { FileIndexStore } from './store'

export interface IndexProgress {
  /** files currently in the index */
  indexed: number
  /** files waiting for extraction */
  pending: number
  scanning: boolean
}

export interface IndexerSources {
  /** the folders walked recursively: the save folder and every added root */
  roots: () => readonly string[]
  /** files outside the roots that should still be searchable (recents, starred) */
  extraPaths: () => readonly string[]
}

const RESCAN_DEBOUNCE_MS = 1500
/**
 * a worker request left unanswered this long means a wedged parse or a stalled
 * walk: the request fails as an extraction error and the worker is recycled,
 * so one bad file cannot stall indexing (and search with it) for good
 */
const WORKER_REQUEST_TIMEOUT_MS = 120_000
/**
 * heap cap for the extraction worker: without resourceLimits a parse that
 * exhausts V8's heap is a process-wide fatal OOM and takes the whole app down
 * with it; with one, only the worker dies and the file is recorded as an error.
 * 256 MB fits the heaviest extraction measured (a 40 MB, 7.6 M-cell xlsx) and
 * keeps V8 collecting instead of letting a scan's garbage pile up to a gigabyte
 */
const WORKER_HEAP_LIMIT_MB = 256
/**
 * The worker is a thread of the main process, so its heap shows up as the
 * app's memory. Once a scan and its extractions are done nothing is lost by
 * retiring it: the next refresh spawns a fresh one in milliseconds, while an
 * idle worker would have held the scan's heap until V8 got around to shrinking it
 */
const WORKER_IDLE_MS = 5_000

/**
 * Keeps the store in step with the disk: a scan diffs mtime/size against the
 * index, changed files queue for extraction on the worker one at a time, and
 * vanished files are dropped. Scans coalesce; extraction is sequential so the
 * user's foreground work keeps the CPU.
 */
export class FileIndexer {
  private worker: Worker | null = null
  private nextId = 1
  private readonly waiting = new Map<number, (r: WorkerResponse) => void>()
  /** request ids posted to each worker: a retired worker's exit fails only its own */
  private readonly postedTo = new WeakMap<Worker, Set<number>>()
  private readonly queue: ScannedFile[] = []
  private readonly queued = new Set<string>()
  private draining = false
  private scanning = false
  private scanRequested = false
  private rescanTimer: NodeJS.Timeout | null = null
  private idleTimer: NodeJS.Timeout | null = null
  private lastScanAt = 0
  private stopped = false

  constructor(
    private readonly store: FileIndexStore,
    private readonly workerPath: string,
    private readonly sources: IndexerSources,
    private readonly requestTimeoutMs = WORKER_REQUEST_TIMEOUT_MS,
    private readonly workerHeapLimitMb = WORKER_HEAP_LIMIT_MB,
    private readonly workerIdleMs = WORKER_IDLE_MS,
  ) {}

  /** whether an extraction worker is currently alive (test seam) */
  workerIsRunning(): boolean {
    return this.worker !== null
  }

  progress(): IndexProgress {
    return { indexed: this.store.count(), pending: this.queue.length, scanning: this.scanning }
  }

  /** schedule a scan soon; repeated calls within the debounce window fold into one */
  refresh(): void {
    if (this.stopped) return
    if (this.rescanTimer) clearTimeout(this.rescanTimer)
    this.rescanTimer = setTimeout(() => {
      this.rescanTimer = null
      void this.scan()
    }, RESCAN_DEBOUNCE_MS)
  }

  /** scan now unless one ran within `maxAgeMs` */
  refreshIfStale(maxAgeMs: number): void {
    if (Date.now() - this.lastScanAt >= maxAgeMs) void this.scan()
  }

  async scan(): Promise<void> {
    if (this.stopped) return
    if (this.scanning) {
      this.scanRequested = true
      return
    }
    this.scanning = true
    try {
      const seen = new Map<string, ScannedFile>()
      let walked = true
      let complete = true
      for (const root of this.sources.roots()) {
        const res = await this.ask({ id: 0, type: 'scan', root })
        // a crashed worker answers with an extract error; dropping the index on that would empty search
        if (res.type === 'scan') {
          for (const f of res.files) seen.set(f.path, f)
          if (res.truncated) complete = false
        } else walked = false
      }
      if (walked) this.diff(seen, complete)
    } finally {
      this.scanning = false
      this.scheduleIdleRecycle()
    }
    // a refresh that arrived mid-scan runs now, whether or not the walk succeeded
    if (this.scanRequested) {
      this.scanRequested = false
      void this.scan()
    }
  }

  /** bring the store in step with what the walk saw; extra paths are stat-ed here */
  // a budget-capped walk cannot tell a vanished file from one it never reached,
  // so entries are only removed when every root was walked to the end
  private diff(seen: Map<string, ScannedFile>, complete: boolean): void {
    for (const p of this.sources.extraPaths()) {
      if (seen.has(p) || !isSupportedTreeFile(p)) continue
      const st = statOrNull(p)
      if (st) seen.set(p, st)
    }
    const known = this.store.listAll()
    if (complete) {
      const gone: string[] = []
      for (const path of known.keys()) if (!seen.has(path)) gone.push(path)
      this.store.remove(gone)
    }
    for (const f of seen.values()) {
      // an unchanged file that already failed is not retried: a parse that
      // kills the worker would otherwise be re-run on every scan
      const k = known.get(f.path)
      if (k && k.mtimeMs === f.mtimeMs && k.sizeBytes === f.sizeBytes) continue
      this.enqueue(f)
    }
    this.lastScanAt = Date.now()
  }

  private enqueue(f: ScannedFile): void {
    if (this.queued.has(f.path)) return
    this.queued.add(f.path)
    this.queue.push(f)
    void this.drain()
  }

  private async drain(): Promise<void> {
    if (this.draining) return
    this.draining = true
    try {
      while (this.queue.length && !this.stopped) {
        const f = this.queue.shift()!
        this.queued.delete(f.path)
        // the file may have changed again while queued; index what is on disk now
        const st = statOrNull(f.path)
        if (!st) {
          this.store.remove([f.path])
          continue
        }
        const res = await this.ask({ id: 0, type: 'extract', path: st.path })
        if (res.type !== 'extract') continue
        this.apply(st, res.result)
      }
    } finally {
      this.draining = false
      this.scheduleIdleRecycle()
    }
  }

  private scheduleIdleRecycle(): void {
    if (this.idleTimer) clearTimeout(this.idleTimer)
    this.idleTimer = setTimeout(() => {
      this.idleTimer = null
      if (this.scanning || this.draining || this.queue.length > 0 || this.waiting.size > 0) return
      this.recycleWorker()
    }, this.workerIdleMs)
    this.idleTimer.unref?.()
  }

  private apply(f: ScannedFile, r: Extracted): void {
    try {
      if (r.kind === 'text') this.store.upsert(f, r.text, 'ok')
      else if (r.kind === 'name-only') this.store.upsert(f, null, 'name-only')
      else this.store.upsert(f, null, 'error')
    } catch {
      // a corrupt row must not stall the queue; the next scan retries it
    }
  }

  private ask(req: WorkerRequest): Promise<WorkerResponse> {
    const id = this.nextId++
    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        // a wedged worker never answers: fail this request the way a crashed
        // worker's requests fail, and retire the worker so later requests get a
        // fresh one instead of hanging too
        this.waiting.delete(id)
        this.recycleWorker()
        resolve({
          id,
          type: 'extract',
          result: { kind: 'error', error: 'worker request timed out' },
        })
      }, this.requestTimeoutMs)
      this.waiting.set(id, (r: WorkerResponse) => {
        clearTimeout(timer)
        resolve(r)
      })
      try {
        const w = this.ensureWorker()
        this.postedTo.get(w)?.add(id)
        w.postMessage({ ...req, id })
      } catch (e) {
        // a failed post must not reject into the void-ed callers: answer as an
        // extraction error, exactly like the drop handler does for a crash
        clearTimeout(timer)
        this.waiting.delete(id)
        resolve({
          id,
          type: 'extract',
          result: { kind: 'error', error: e instanceof Error ? e.message : String(e) },
        })
      }
    })
  }

  /** retire the current worker; 'exit' fails any other in-flight request via drop */
  private recycleWorker(): void {
    const w = this.worker
    if (!w) return
    this.worker = null
    void w.terminate()
  }

  private ensureWorker(): Worker {
    if (this.idleTimer) {
      clearTimeout(this.idleTimer)
      this.idleTimer = null
    }
    if (this.worker) return this.worker
    const w = new Worker(this.workerPath, {
      resourceLimits: { maxOldGenerationSizeMb: this.workerHeapLimitMb },
    })
    w.on('message', (msg: WorkerResponse) => {
      const cb = this.waiting.get(msg.id)
      if (!cb) return
      this.waiting.delete(msg.id)
      cb(msg)
    })
    const drop = () => {
      // a crashed worker fails its in-flight request as an extraction error so
      // the queue moves on; the exit of a worker retired for idleness or a
      // timeout arrives after its successor may already have requests in
      // flight, so only this worker's own requests are failed
      if (this.worker === w) this.worker = null
      const posted = this.postedTo.get(w) ?? new Set<number>()
      for (const id of posted) {
        const cb = this.waiting.get(id)
        if (!cb) continue
        this.waiting.delete(id)
        cb({ id, type: 'extract', result: { kind: 'error', error: 'worker exited' } })
      }
      posted.clear()
    }
    this.postedTo.set(w, new Set())
    w.on('error', drop)
    w.on('exit', drop)
    this.worker = w
    return w
  }

  stop(): void {
    this.stopped = true
    if (this.rescanTimer) clearTimeout(this.rescanTimer)
    if (this.idleTimer) clearTimeout(this.idleTimer)
    this.queue.length = 0
    this.queued.clear()
    void this.worker?.terminate()
    this.worker = null
  }
}
