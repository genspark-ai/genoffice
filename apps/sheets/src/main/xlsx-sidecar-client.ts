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
}

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

  constructor(private readonly binaryPath: string) {}

  async open(path: string, locale = 'zh'): Promise<unknown> {
    return this.request({ command: 'open', path, locale })
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

  async close(sessionId: string): Promise<void> {
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

  /** spawn the sidecar ahead of the first request to hide process cold-start */
  start(): void {
    this.ensureStarted()
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
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(requestId)
        reject(new Error('XLSX sidecar request timed out.'))
      }, timeoutMs)
      this.pending.set(requestId, { resolve, reject, timeout })
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
    const child = spawn(this.binaryPath, [], {
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
    })
    this.process = child
    this.stderr = ''
    this.lines = createInterface({ input: child.stdout })
    this.lines.on('line', (line) => this.handleLine(line))
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => {
      this.stderr = `${this.stderr}${chunk}`.slice(-MAX_STDERR_LENGTH)
    })
    child.once('error', (error) => {
      this.process = null
      this.rejectPending(error)
    })
    child.once('exit', (code, signal) => {
      this.process = null
      this.lines?.close()
      this.lines = null
      const detail = this.stderr.trim()
      const reason = detail
        ? `XLSX sidecar exited: ${detail}`
        : `XLSX sidecar exited with code ${String(code)} and signal ${String(signal)}.`
      this.rejectPending(new Error(reason))
    })
    return child
  }

  private handleLine(line: string): void {
    let response: SidecarResponse
    try {
      response = JSON.parse(line) as SidecarResponse
    } catch {
      this.rejectPending(new Error('XLSX sidecar returned invalid JSON.'))
      return
    }
    if (
      response.version !== PROTOCOL_VERSION ||
      typeof response.requestId !== 'string' ||
      typeof response.ok !== 'boolean'
    ) {
      this.rejectPending(new Error('XLSX sidecar returned an invalid response.'))
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
