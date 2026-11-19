import { EventEmitter } from 'node:events'
import { PassThrough } from 'node:stream'

import { afterEach, describe, expect, it, vi } from 'vitest'

class FakeSidecarProcess extends EventEmitter {
  readonly stdin = new PassThrough()
  readonly stdout = new PassThrough()
  readonly stderr = new PassThrough()
  readonly pid: number
  killed = false
  constructor(pid: number) {
    super()
    this.pid = pid
  }
  kill(): void {
    this.killed = true
  }
}

const spawnMock = vi.hoisted(() => vi.fn())
vi.mock('node:child_process', () => ({ spawn: spawnMock }))

const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }

describe('XlsxSidecarClient restart', () => {
  afterEach(() => spawnMock.mockReset())

  it('a late exit from the stopped child leaves its replacement serving', async () => {
    const first = new FakeSidecarProcess(101)
    const second = new FakeSidecarProcess(102)
    spawnMock.mockReturnValueOnce(first).mockReturnValueOnce(second)
    const { XlsxSidecarClient } = await import('../src/main/xlsx-sidecar-client')
    const client = new XlsxSidecarClient('/nonexistent/sidecar')

    const before = client.readRange({ sessionId: 's-1', sheetId: 'sheet-1', range })
    client.stop()
    await expect(before).rejects.toThrow('XLSX sidecar stopped.')

    const after = client.readRange({ sessionId: 's-1', sheetId: 'sheet-1', range })
    expect(client.getProcessId()).toBe(102)
    first.emit('exit', 0, null)
    expect(client.getProcessId()).toBe(102)

    const raw = second.stdin.read() as Buffer
    const request = JSON.parse(raw.toString('utf8').trim()) as { requestId: string }
    second.stdout.write(
      `${JSON.stringify({ version: 1, requestId: request.requestId, ok: true, result: { cells: [] } })}\n`,
    )
    await expect(after).resolves.toEqual({ cells: [] })
  })
})

describe('XlsxSidecarClient error teardown', () => {
  it('an error closes the reader so the following exit has nothing to leak', async () => {
    const fake = new FakeSidecarProcess(201)
    spawnMock.mockReturnValue(fake)
    const { XlsxSidecarClient } = await import('../src/main/xlsx-sidecar-client')
    const client = new XlsxSidecarClient('/nonexistent/sidecar')
    const read = client.readRange({ sessionId: 's-1', sheetId: 'sheet-1', range })
    expect(fake.stdout.listenerCount('data')).toBeGreaterThan(0)
    fake.emit('error', new Error('spawn failed'))
    await expect(read).rejects.toThrow('spawn failed')
    expect(fake.stdout.listenerCount('data')).toBe(0)
    fake.emit('exit', 1, null)
    expect(client.getProcessId()).toBeNull()
  })
})

describe('XlsxSidecarClient stdin error', () => {
  afterEach(() => spawnMock.mockReset())

  it('an async stdin EPIPE tears the client down instead of crashing the app', async () => {
    const child = new FakeSidecarProcess(201)
    spawnMock.mockReturnValueOnce(child)
    const { XlsxSidecarClient } = await import('../src/main/xlsx-sidecar-client')
    const client = new XlsxSidecarClient('/nonexistent/sidecar')

    // a pending request, then the child dies OOM-style: killed stays false
    const pending = client.readRange({ sessionId: 's', sheetId: 'sheet', range })
    // PassThrough with no consumer end: destroy() emits the async 'error'
    child.stdin.destroy(new Error('write EPIPE'))

    await expect(pending).rejects.toThrow(/stdin failed|EPIPE/)
    const again = client.readRange({ sessionId: 's', sheetId: 'sheet', range })
    expect(spawnMock.mock.calls.length).toBeGreaterThan(1)
    void again.catch(() => {})
    child.stdout.end()
    child.stderr.end()
    client.stop()
  })
})

/// The renderer recovers a crashed workbook from this signal, so it must be
/// raised by an actual process death and by nothing else. In particular the
/// Save swap and closeWorkbook tear a session down on purpose while the
/// sidecar keeps running, and must never look like a crash.
describe('XlsxSidecarClient crash signal', () => {
  afterEach(() => spawnMock.mockReset())

  it('reports a process death, which invalidates every session id we hold', async () => {
    const fake = new FakeSidecarProcess(301)
    spawnMock.mockReturnValue(fake)
    const { XlsxSidecarClient } = await import('../src/main/xlsx-sidecar-client')
    const client = new XlsxSidecarClient('/nonexistent/sidecar')
    const crashes = vi.fn()
    client.onProcessExit(crashes)

    const read = client.readRange({ sessionId: 's', sheetId: 'sheet', range })
    fake.emit('exit', 1, null)
    await expect(read).rejects.toThrow(/exited/)
    expect(crashes).toHaveBeenCalledTimes(1)
    // The signal fires once per death, not once per rejected request: a crash
    // fails every in-flight read at once and each must not re-open the file.
    void client.readRange({ sessionId: 's', sheetId: 'sheet', range }).catch(() => {})
    expect(crashes).toHaveBeenCalledTimes(1)
    client.stop()
  })

  it('stays silent for the deliberate close of one workbook', async () => {
    const fake = new FakeSidecarProcess(302)
    spawnMock.mockReturnValue(fake)
    const { XlsxSidecarClient } = await import('../src/main/xlsx-sidecar-client')
    const client = new XlsxSidecarClient('/nonexistent/sidecar')
    const crashes = vi.fn()
    client.onProcessExit(crashes)

    // What closeWorkbook and the Save swap do: drop the session through the
    // live pipe. The process is untouched, so the renderer must not be told
    // the workbook needs re-opening.
    const closing = client.close('s-1')
    const closeRequest = fake.stdin.read() as Buffer
    fake.stdout.write(
      `${JSON.stringify({
        version: 1,
        requestId: (JSON.parse(closeRequest.toString('utf8').trim()) as { requestId: string })
          .requestId,
        ok: true,
        result: null,
      })}\n`,
    )
    await closing

    expect(crashes).not.toHaveBeenCalled()
    expect(client.getProcessId()).toBe(302)
    client.stop()
  })

  it('stays silent for the app-quit stop', async () => {
    const fake = new FakeSidecarProcess(303)
    spawnMock.mockReturnValue(fake)
    const { XlsxSidecarClient } = await import('../src/main/xlsx-sidecar-client')
    const client = new XlsxSidecarClient('/nonexistent/sidecar')
    const crashes = vi.fn()
    client.onProcessExit(crashes)

    client.start()
    client.stop()
    // The kill is deliberate; the late exit event it triggers must not read
    // as a crash either.
    fake.emit('exit', 0, null)
    expect(crashes).not.toHaveBeenCalled()
  })

  it('unsubscribes cleanly', async () => {
    const fake = new FakeSidecarProcess(304)
    spawnMock.mockReturnValue(fake)
    const { XlsxSidecarClient } = await import('../src/main/xlsx-sidecar-client')
    const client = new XlsxSidecarClient('/nonexistent/sidecar')
    const crashes = vi.fn()
    const off = client.onProcessExit(crashes)

    client.start()
    off()
    fake.emit('exit', 1, null)
    expect(crashes).not.toHaveBeenCalled()
    client.stop()
  })
})
