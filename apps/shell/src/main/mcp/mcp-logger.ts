import {
  closeSync,
  existsSync,
  openSync,
  readFileSync,
  readSync,
  statSync,
  appendFileSync,
  truncateSync,
  writeFileSync,
} from 'node:fs'

/**
 * Local log for the MCP server: a bounded in-memory ring plus an append-only
 * file under userData, so both the settings UI (recent lines) and a saved file
 * the user can open/share are covered. Appending is synchronous on purpose —
 * log volume is tiny (lifecycle, sessions, tool calls) and a lost line on
 * crash is acceptable, while ordering guarantees are not.
 */

const MAX_BUFFER_LINES = 500
/** how many lines the settings pane fetches at once */
export const MCP_LOG_TAIL = 200

/** Device-local timestamp for log lines (YYYY-MM-DD HH:mm:ss.SSS): the log is
 *  read by the user in the settings pane / a shared file, so it shows wall-clock
 *  time, not UTC (toISOString's Z suffix read as a 8h-off time in zh locales). */
export function localTimestamp(date = new Date()): string {
  const pad = (value: number, width = 2): string => String(value).padStart(width, '0')
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ` +
    `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}.` +
    `${pad(date.getMilliseconds(), 3)}`
  )
}

export class McpLogger {
  private readonly buffer: string[] = []

  constructor(readonly filePath: string) {
    // the log covers the current app launch only: reset the file at startup so
    // it never grows across runs (the settings pane reads it live anyway)
    try {
      writeFileSync(this.filePath, '', 'utf8')
    } catch {
      // best-effort; append() recreates what it can
    }
  }

  append(message: string): void {
    const line = `[${localTimestamp()}] ${message}`
    this.buffer.push(line)
    if (this.buffer.length > MAX_BUFFER_LINES) this.buffer.shift()
    try {
      appendFileSync(this.filePath, line + '\n', 'utf8')
    } catch {
      // a broken log target must never take the MCP server down
    }
  }

  /** most recent lines, file-backed (falls back to the ring when unreadable) */
  recent(limit = MCP_LOG_TAIL): string[] {
    try {
      if (existsSync(this.filePath)) {
        const raw = readTail(this.filePath, limit)
        if (raw.length > 0) return raw
      }
    } catch {
      // fall through to the ring buffer
    }
    return this.buffer.slice(-limit)
  }

  clear(): void {
    this.buffer.length = 0
    try {
      if (existsSync(this.filePath)) truncateSync(this.filePath, 0)
    } catch {
      // best-effort
    }
  }

  /** create the file when missing so "reveal in file manager" has a target */
  ensureFile(): void {
    try {
      if (!existsSync(this.filePath)) writeFileSync(this.filePath, '', 'utf8')
    } catch {
      // best-effort
    }
  }
}

/**
 * Last `limit` non-empty lines of a text file.
 *
 * The settings pane polls this every couple of seconds while logging is on, so
 * it reads a bounded suffix instead of the whole append-only file: growing the
 * work with total history would put an unbounded synchronous read on the main
 * process. A generous per-line allowance covers the long lines a stack trace
 * can produce; falling back to the whole file when the suffix is too small to
 * hold `limit` lines keeps the result exact.
 */
function readTail(filePath: string, limit: number): string[] {
  const MAX_BYTES = 256 * 1024
  const { size } = statSync(filePath)
  if (size <= MAX_BYTES) return tailLines(readFileSync(filePath, 'utf8'), limit)

  const start = size - MAX_BYTES
  const fh = openSync(filePath, 'r')
  try {
    const buffer = Buffer.allocUnsafe(MAX_BYTES)
    const read = readSync(fh, buffer, 0, MAX_BYTES, start)
    let lines = buffer.subarray(0, read).toString('utf8').split('\n')
    // the window starts mid-line, so its first entry is a partial line
    lines = lines.slice(1)
    return lines.filter((l) => l.length > 0).slice(-limit)
  } finally {
    closeSync(fh)
  }
}

function tailLines(raw: string, limit: number): string[] {
  return raw
    .split('\n')
    .filter((l) => l.length > 0)
    .slice(-limit)
}
