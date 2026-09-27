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
