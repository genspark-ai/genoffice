import { EventEmitter } from 'node:events'
import { PassThrough } from 'node:stream'

import { afterEach, describe, expect, it, vi } from 'vitest'

class FakeSidecarProcess extends EventEmitter {
  readonly stdin = new PassThrough()
  readonly stdout = new PassThrough()
  readonly stderr = new PassThrough()
  readonly killed = false
  kill(): void {}
}

function writtenRequests(fake: FakeSidecarProcess): Record<string, unknown>[] {
  const raw = fake.stdin.read() as Buffer | null
  if (!raw) return []
  return raw
    .toString('utf8')
    .split('\n')
    .filter((line) => line.length > 0)
    .map((line) => JSON.parse(line) as Record<string, unknown>)
}

const spawnMock = vi.hoisted(() => vi.fn())
vi.mock('node:child_process', () => ({ spawn: spawnMock }))

describe('XlsxSidecarClient cancel wiring', () => {
  afterEach(() => {
    vi.useRealTimers()
    spawnMock.mockReset()
  })

  async function clientWithFake() {
    const fake = new FakeSidecarProcess()
    spawnMock.mockReturnValue(fake)
    const { XlsxSidecarClient } = await import('../src/main/xlsx-sidecar-client')
    return { client: new XlsxSidecarClient('/nonexistent/sidecar'), fake }
  }

  it('close sweeps the session queued reads with cancels ahead of the close', async () => {
    const { client, fake } = await clientWithFake()
    const readA = client.readRange({
      sessionId: 's-1',
      sheetId: 'sheet-1',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    })
    const readOther = client.readRange({
      sessionId: 's-2',
      sheetId: 'sheet-1',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    })
    const closed = client.close('s-1')
    const requests = writtenRequests(fake)
    expect(requests.map((request) => request.command)).toEqual([
      'read_range',
      'read_range',
      'cancel',
      'close',
    ])
    const cancel = requests[2]!
    expect(cancel.targetRequestId).toBe(requests[0]!.requestId)

    fake.stdout.write(
      `${JSON.stringify({
        version: 1,
        requestId: requests[0]!.requestId,
        ok: false,
        error: { code: 'cancelled', message: 'Request was cancelled by the client.' },
      })}\n`,
    )
    await expect(readA).rejects.toThrow('Request was cancelled by the client.')
    fake.stdout.write(
      `${JSON.stringify({ version: 1, requestId: requests[3]!.requestId, ok: true, result: { closed: true } })}\n`,
    )
    await expect(closed).resolves.toBeUndefined()

    fake.stdout.write(
      `${JSON.stringify({ version: 1, requestId: requests[1]!.requestId, ok: true, result: { cells: [] } })}\n`,
    )
    await expect(readOther).resolves.toEqual({ cells: [] })
  })

  it('a timed-out close is never cancelled: skipping it would leak the session', async () => {
    vi.useFakeTimers()
    const { client, fake } = await clientWithFake()
    const closed = client.close('s-1')
    vi.advanceTimersByTime(30_000)
    await expect(closed).rejects.toThrow('timed out')
    expect(writtenRequests(fake).map((request) => request.command)).toEqual(['close'])
  })

  it('a client-side timeout sends a cancel for the abandoned request', async () => {
    vi.useFakeTimers()
    const { client, fake } = await clientWithFake()
    const read = client.readRange({
      sessionId: 's-1',
      sheetId: 'sheet-1',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    })
    const [request] = writtenRequests(fake)
    vi.advanceTimersByTime(30_000)
    await expect(read).rejects.toThrow('timed out')
    const [cancel] = writtenRequests(fake)
    expect(cancel?.command).toBe('cancel')
    expect(cancel?.targetRequestId).toBe(request!.requestId)
  })
})
