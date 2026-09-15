import { existsSync, readFileSync } from 'node:fs'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { McpLogger } from '../../src/main/mcp/mcp-logger'

/** the MCP server's local log: bounded ring + append-only file under userData */

let dir: string
let logPath: string
let logger: McpLogger

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-log-'))
  logPath = join(dir, 'mcp-log.txt')
  logger = new McpLogger(logPath)
})

afterEach(async () => {
  await rm(dir, { recursive: true, force: true })
})

describe('McpLogger', () => {
  it('appends timestamped lines to the file and returns them from recent()', () => {
    logger.append('[mcp] listening on http://127.0.0.1:3093')
    logger.append('[mcp] tool create_docx ok (12ms)')

    const lines = logger.recent()
    expect(lines).toHaveLength(2)
    expect(lines[1]).toContain('[mcp] tool create_docx ok (12ms)')
    // local-time timestamp prefix (YYYY-MM-DD HH:mm:ss.SSS)
    expect(lines[0]).toMatch(/^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3}\]/)

    const onDisk = readFileSync(logPath, 'utf8').split('\n').filter(Boolean)
    expect(onDisk).toHaveLength(2)
  })

  it('timestamps are device-local wall-clock time, not UTC', () => {
    logger.append('tz check')
    const line = logger.recent()[0] ?? ''
    const m = /^\[(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})\.(\d{3})\]/.exec(line)
    expect(m, line).not.toBeNull()
    // interpreting the stamp in the device timezone lands ~now; a UTC ISO
    // stamp (the old format) drifts by the timezone offset here
    const parsed = new Date(
      Number(m![1]),
      Number(m![2]) - 1,
      Number(m![3]),
      Number(m![4]),
      Number(m![5]),
      Number(m![6]),
      Number(m![7]),
    )
    expect(Math.abs(parsed.getTime() - Date.now())).toBeLessThan(5_000)
  })

  it('recent() prefers the file contents and caps the tail', () => {
    for (let i = 0; i < 30; i++) logger.append(`line ${i}`)
    const tail = logger.recent(10)
    expect(tail).toHaveLength(10)
    expect(tail[0]).toContain('line 20')
    expect(tail[9]).toContain('line 29')
  })

  it('recent() reads the tail of a file larger than the suffix window', () => {
    // write past the 256 KiB suffix window so the bounded-read path runs; the
    // last lines must still come back whole (no truncated first line)
    const big = 'x'.repeat(4096)
    for (let i = 0; i < 100; i++) logger.append(`line ${i} ${big}`)
    const tail = logger.recent(5)
    expect(tail).toHaveLength(5)
    expect(tail[0]).toContain('line 95 ')
    expect(tail[4]).toContain('line 99 ')
    // every returned line is complete: the partial window edge is dropped
    for (const line of tail) expect(line).toMatch(/^\[\d{4}-.*\] line \d+ x+$/)
  })

  it('starts each logger with a fresh file (per-launch log, not persisted across runs)', () => {
    logger.append('[mcp] line from a previous run')
    // a new app launch constructs a fresh logger over the same path: the old
    // content is gone, the new run's lines start from zero
    const fresh = new McpLogger(logPath)
    expect(fresh.recent()).toEqual([])
    expect(readFileSync(logPath, 'utf8')).toBe('')
    fresh.append('[mcp] new run')
    const lines = fresh.recent()
    expect(lines).toHaveLength(1)
    expect(lines[0]).toContain('[mcp] new run')
  })

  it('clear() truncates the file and the ring', () => {
    logger.append('one')
    logger.clear()
    expect(logger.recent()).toEqual([])
    expect(existsSync(logPath)).toBe(true)
    expect(readFileSync(logPath, 'utf8')).toBe('')
  })

  it('the file exists from construction on (ensureFile is a safe no-op)', () => {
    // the constructor resets the file, so "reveal in file manager" always has
    // a target without calling ensureFile first
    expect(existsSync(logPath)).toBe(true)
    logger.ensureFile()
    expect(readFileSync(logPath, 'utf8')).toBe('')
  })
})
