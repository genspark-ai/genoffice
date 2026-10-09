import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { createInterface, type Interface } from 'node:readline'

const PROTOCOL_VERSION = 1
const REQUEST_TIMEOUT_MS = 30_000
// Archive commands stream whole workbooks; large files need more headroom.
const ARCHIVE_TIMEOUT_MS = 180_000
const MAX_STDERR_LENGTH = 8_192

interface PendingRequest {
  readonly resolve: (value: unknown) => void
  readonly reject: (error: Error) => void
  readonly timeout: NodeJS.Timeout
  readonly command: string
  /** Session-scoped commands record it so close() can sweep their queue slots. */
  readonly sessionId?: string
}

/** Pure reads a departing consumer can abandon; skipping a queued close (or
 * save) instead would leave its side effects permanently undone. */
const CANCELLABLE_READS = new Set([
  'read_range',
  'read_formula_cells',
  'read_row_outline',
  'read_media',
  'find_cells',
])

interface SidecarResponse {
  readonly version: number
  readonly requestId: string
  readonly ok: boolean
  readonly result?: unknown
  readonly error?: {
    readonly code: string
    readonly message: string
  }
}

export class XlsxSidecarClient {
  private process: ChildProcessWithoutNullStreams | null = null
  private lines: Interface | null = null
  private readonly pending = new Map<string, PendingRequest>()
  private stderr = ''
  /// Notified when the sidecar process dies and the NEXT request will spawn a
  /// replacement that has never heard of the session ids we are holding. This
  /// is a positive crash signal, and deliberately NOT the same thing as a
  /// session guard rejecting an id: closing a workbook (or the Save swap that
  /// replaces its session) only sends a `close` command down the live pipe, so
  /// it never reaches here. A renderer that keyed recovery off a guard's error
  /// text could not tell those two apart, and would re-open the file on every
  /// ordinary save.
  private readonly exitListeners = new Set<() => void>()

  constructor(private readonly binaryPath: string) {}

  /** Subscribe to unexpected sidecar process death; returns an unsubscribe. */
  onProcessExit(listener: () => void): () => void {
    this.exitListeners.add(listener)
    return () => this.exitListeners.delete(listener)
  }

  private notifyExit(): void {
    for (const listener of this.exitListeners) listener()
  }

  async open(path: string, locale = 'zh', shortDateFormat?: string): Promise<unknown> {
    return this.request({
      command: 'open',
      path,
      locale,
      ...(shortDateFormat === undefined ? {} : { shortDateFormat }),
    })
  }

  async readRange(input: {
    readonly sessionId: string
    readonly sheetId: string
    readonly range: {
      readonly startRow: number
      readonly endRow: number
      readonly startColumn: number
      readonly endColumn: number
    }
  }): Promise<unknown> {
    return this.request({ command: 'read_range', ...input })
  }

  async readFormulaCells(input: {
    readonly sessionId: string
    readonly sheetId: string
  }): Promise<unknown> {
    return this.request({ command: 'read_formula_cells', ...input })
  }

  async findCells(
    input: Readonly<Record<string, unknown>> & { readonly sessionId: string },
  ): Promise<unknown> {
    return this.request({ command: 'find_cells', ...input })
  }

  async readRowOutline(input: {
    readonly sessionId: string
    readonly sheetId: string
  }): Promise<unknown> {
    return this.request({ command: 'read_row_outline', ...input })
  }

  async close(sessionId: string): Promise<void> {
    // Skip the session's queued reads so close does not wait behind them;
    // each settles with the sidecar's authoritative "cancelled" failure.
    for (const [requestId, pending] of this.pending) {
      if (pending.sessionId === sessionId && CANCELLABLE_READS.has(pending.command)) {
        this.sendCancel(requestId)
      }
    }
    await this.request({ command: 'close', sessionId })
  }

  async readMedia(input: {
    readonly sessionId: string
    readonly visualId: string
  }): Promise<unknown> {
    return this.request({ command: 'read_media', ...input })
  }

  async convertWorkbook(input: { path: string; targetPath: string }): Promise<unknown> {
    return this.request({ command: 'convert_workbook', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async archiveManifest(path: string): Promise<unknown> {
    return this.request({ command: 'archive_manifest', path }, ARCHIVE_TIMEOUT_MS)
  }

  async readEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly outputDir: string
  }): Promise<unknown> {
    return this.request({ command: 'read_entries', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async scanEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly needle: string
  }): Promise<unknown> {
    return this.request({ command: 'scan_entries', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async saveArchive(input: {
    readonly sourcePath: string
    readonly targetPath: string
    readonly replacements: readonly { name: string; contentPath: string }[]
    readonly removals: readonly string[]
    readonly additions: readonly { name: string; contentPath: string }[]
  }): Promise<unknown> {
    return this.request({ command: 'save_archive', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  /** IronCalc-backed recalculation: load the file, apply the pending edits,
   * evaluate, and return the requested cells' computed values. Fail-soft —
   * callers keep cached values when this errors. */
  async recalcCells(input: {
    readonly path: string
    readonly edits: readonly {
      readonly sheet: string
      readonly row: number
      readonly column: number
      readonly input: string
    }[]
    readonly reads: readonly {
      readonly sheet: string
      readonly range: {
        readonly startRow: number
        readonly endRow: number
        readonly startColumn: number
        readonly endColumn: number
      }
    }[]
  }): Promise<unknown> {
    return this.request({ command: 'recalc_cells', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  /**
   * Spawn the sidecar ahead of the first request to hide process cold-start.
   * The prewarm is best-effort: on Windows an unrunnable binary (corrupt,
   * AV-blocked — CreateProcess errors outside Node's delayed-error set, e.g.
   * UNKNOWN) makes spawn() throw synchronously, and letting that propagate
   * crashed tab/window creation with an uncaught main-process exception. The
   * first real request retries the spawn and surfaces the descriptive error
   * through its own handled rejection path instead.
   */
  start(): void {
    try {
      this.ensureStarted()
    } catch (error) {
      console.warn('[sheets] XLSX sidecar prewarm failed:', error)
    }
  }

  getProcessId(): number | null {
    return this.process?.pid ?? null
  }

  stop(): void {
    this.lines?.close()
    this.lines = null
    this.process?.kill()
    this.process = null
    this.rejectPending(new Error('XLSX sidecar stopped.'))
  }

  /**
   * Fire-and-forget: handled out of band by the sidecar's reader thread, so
   * it takes effect while earlier requests still wait in the queue. The
   * reply matches no pending entry and is dropped.
   */
  private sendCancel(targetRequestId: string): void {
    const child = this.process
    if (!child || child.killed) return
    child.stdin.write(
      `${JSON.stringify({
        version: PROTOCOL_VERSION,
        requestId: crypto.randomUUID(),
        command: 'cancel',
        targetRequestId,
      })}\n`,
    )
  }

  private request(
    command: Readonly<Record<string, unknown>>,
    timeoutMs: number = REQUEST_TIMEOUT_MS,
  ): Promise<unknown> {
    const child = this.ensureStarted()
    const requestId = crypto.randomUUID()
    const payload = JSON.stringify({
      version: PROTOCOL_VERSION,
      requestId,
      ...command,
    })
    const commandName = command.command
    const sessionId = typeof command.sessionId === 'string' ? command.sessionId : undefined
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(requestId)
        // Still queued sidecar-side, the request would execute for a reply
        // nobody consumes (a 100k-cell range serializes ~5MB of JSON). A
        // close is the exception: its side effect IS the cleanup, and no
        // caller retries a timed-out close — skipping it would leak the
        // session's index threads and cache for good.
        if (commandName !== 'close') this.sendCancel(requestId)
        reject(new Error('XLSX sidecar request timed out.'))
      }, timeoutMs)
      this.pending.set(requestId, {
        resolve,
        reject,
        timeout,
        command: String(commandName),
        ...(sessionId === undefined ? {} : { sessionId }),
      })
      child.stdin.write(`${payload}\n`, (error) => {
        if (!error) return
        const pending = this.pending.get(requestId)
        if (!pending) return
        clearTimeout(pending.timeout)
        this.pending.delete(requestId)
        pending.reject(error)
      })
    })
  }

  private ensureStarted(): ChildProcessWithoutNullStreams {
    if (this.process && !this.process.killed) return this.process
    let child: ChildProcessWithoutNullStreams
    try {
      child = spawn(this.binaryPath, [], {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
      })
    } catch (error) {
      // Synchronous spawn failures ("spawn UNKNOWN") carry no path in the
      // message — rethrow with enough context to diagnose from a screenshot.
      throw new Error(
        `XLSX sidecar failed to start (${this.binaryPath}): ${error instanceof Error ? error.message : String(error)}`,
        { cause: error },
      )
    }
    this.process = child
    this.stderr = ''
    this.lines = createInterface({ input: child.stdout })
    this.lines.on('line', (line) => this.handleLine(line))
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => {
      this.stderr = `${this.stderr}${chunk}`.slice(-MAX_STDERR_LENGTH)
    })
    const teardown = (reason: Error): void => {
      if (this.process !== child) return
      this.process = null
      this.lines?.close()
      this.lines = null
      this.rejectPending(reason)
      // The process is gone, so every session id we hold is now unknown to
      // the next spawn. stop() (app quit) sets this.process to null first,
      // which makes the guard above return, so a deliberate shutdown never
      // reaches this and never reads as a crash.
      this.notifyExit()
    }
    child.once('error', teardown)
    // An externally killed child keeps `killed` false, so the next write emits
    // an async EPIPE that would otherwise be an uncaught exception.
    child.stdin.on('error', () => {
      teardown(
        new Error(
          `XLSX sidecar stdin failed (${this.stderr.trim() || 'broken pipe'}).`.slice(0, 200),
        ),
      )
    })
    child.once('exit', (code, signal) => {
      const detail = this.stderr.trim()
      const reason = detail
        ? `XLSX sidecar exited: ${detail}`
        : `XLSX sidecar exited with code ${String(code)} and signal ${String(signal)}.`
      teardown(new Error(reason))
    })
    return child
  }

  private handleLine(line: string): void {
    let response: SidecarResponse
    try {
      response = JSON.parse(line) as SidecarResponse
    } catch {
      // Library code inside the sidecar can print diagnostics to stdout
      // (IronCalc's importer does, e.g. "Unexpected type (empty) in ...").
      // Treat such lines as noise: rejecting every pending request here
      // nuked all in-flight chunk loads whenever a recalc hit one of those
      // workbooks. A genuinely garbled response surfaces as that one
      // request's timeout instead.
      this.stderr = `${this.stderr}[stdout] ${line}\n`.slice(-MAX_STDERR_LENGTH)
      return
    }
    if (
      response.version !== PROTOCOL_VERSION ||
      typeof response.requestId !== 'string' ||
      typeof response.ok !== 'boolean'
    ) {
      this.stderr = `${this.stderr}[stdout] ${line}\n`.slice(-MAX_STDERR_LENGTH)
      return
    }
    const pending = this.pending.get(response.requestId)
    if (!pending) return
    clearTimeout(pending.timeout)
    this.pending.delete(response.requestId)
    if (response.ok) {
      pending.resolve(response.result)
      return
    }
    pending.reject(new Error(response.error?.message ?? 'XLSX sidecar request failed.'))
  }

  private rejectPending(error: Error): void {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timeout)
      pending.reject(error)
    }
    this.pending.clear()
  }
}
