import { EventEmitter } from 'node:events'
import { PassThrough } from 'node:stream'

import { describe, expect, it, vi } from 'vitest'
import type { spawn } from 'node:child_process'

import {
  lpArgumentsFor,
  lpPageRange,
  printFileViaCups,
  printPdfViaCups,
  printToPdfOptionsFor,
} from '../src/print-cups'
import { printOutcome } from '../src/print-job'

/** A stand-in for the lp child process: records the spawn arguments and lets
 * the test drive stdout/stderr/exit the way the real process would. */
class FakeLp extends EventEmitter {
  readonly stdin = new PassThrough()
  readonly stdout = new PassThrough()
  readonly stderr = new PassThrough()
  close(code: number): void {
    queueMicrotask(() => this.emit('close', code))
  }
}

type SpawnCall = { path: string; args: string[]; fake: FakeLp }

function fakeSpawn(): {
  calls: SpawnCall[]
  spawnFn: (path: string, args: string[]) => FakeLp
} {
  const calls: SpawnCall[] = []
  const spawnFn = (path: string, args: string[]) => {
    const fake = new FakeLp()
    calls.push({ path, args, fake })
    // defer so the caller can attach its listeners first
    queueMicrotask(() => fake.emit('spawn'))
    return fake
  }
  return { calls, spawnFn }
}

describe('lpArgumentsFor', () => {
  it('maps the dialog job onto lp flags and omits what was not chosen', () => {
    expect(lpArgumentsFor({})).toEqual([])
    expect(lpArgumentsFor({ deviceName: 'HP' })).toEqual(['-d', 'HP'])
    expect(
      lpArgumentsFor({
        deviceName: 'HP',
        copies: 3,
        collate: true,
        duplexMode: 'longEdge',
      }),
    ).toEqual(['-d', 'HP', '-n', '3', '-o', 'Collate=True', '-o', 'sides=two-sided-long-edge'])
    expect(lpArgumentsFor({ collate: false, duplexMode: 'simplex' })).toEqual([
      '-o',
      'Collate=False',
      '-o',
      'sides=one-sided',
    ])
    expect(lpArgumentsFor({ duplexMode: 'shortEdge' })).toEqual([
      '-o',
      'sides=two-sided-short-edge',
    ])
  })
})

describe('printToPdfOptionsFor', () => {
  it('passes the geometry and drops the defaults', () => {
    expect(printToPdfOptionsFor({})).toEqual({ printBackground: true })
    expect(
      printToPdfOptionsFor({
        widthIn: 8.27,
        heightIn: 11.69,
        marginsIn: { top: 0, bottom: 0, left: 0, right: 0 },
        scale: 1.5,
        landscape: true,
      }),
    ).toEqual({
      printBackground: true,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      pageSize: { width: 8.27, height: 11.69 },
      landscape: true,
      scale: 1.5,
    })
    // scale 1 is Chromium's default: sending it would be noise
    expect(printToPdfOptionsFor({ scale: 1 })).toEqual({ printBackground: true })
  })
})

describe('printPdfViaCups', () => {
  it('spools the pdf on stdin with the job flags and reports the accepted job', async () => {
    const { calls, spawnFn } = fakeSpawn()
    const pending = printPdfViaCups(
      new Uint8Array([1, 2, 3]),
      { deviceName: 'HP', copies: 2 },
      spawnFn as unknown as typeof spawn,
    )
    await vi.waitFor(() => expect(calls).toHaveLength(1))
    expect(calls[0]!.path).toBe('/usr/bin/lp')
    expect(calls[0]!.args).toEqual(['-d', 'HP', '-n', '2'])
    const { fake } = calls[0]!
    fake.stdout.write('request id is HP-42 (1 file(s))\n')
    const written = fake.stdin.read() as Buffer
    expect([...written]).toEqual([1, 2, 3])
    fake.close(0)
    await expect(pending).resolves.toEqual({ ok: true })
  })

  it('surfaces lp stderr when the queue is refused instead of vanishing', async () => {
    const { calls, spawnFn } = fakeSpawn()
    const pending = printPdfViaCups(
      new Uint8Array([1]),
      { deviceName: 'mangled_queue' },
      spawnFn as unknown as typeof spawn,
    )
    await vi.waitFor(() => expect(calls).toHaveLength(1))
    calls[0]!.fake.stderr.write('lp: Invalid destination name.\n')
    calls[0]!.fake.close(1)
    await expect(pending).resolves.toEqual({ ok: false, error: 'lp: Invalid destination name.' })
  })

  it('reports a spawn failure (no lp binary) as an error, not a crash', async () => {
    const spawnFn = () => {
      const fake = new FakeLp()
      queueMicrotask(() => fake.emit('error', new Error('spawn /usr/bin/lp ENOENT')))
      return fake
    }
    await expect(
      printPdfViaCups(new Uint8Array([1]), {}, spawnFn as unknown as typeof spawn),
    ).resolves.toEqual({
      ok: false,
      error: 'spawn /usr/bin/lp ENOENT',
    })
  })

  it('falls back to the exit code when lp says nothing', async () => {
    const { calls, spawnFn } = fakeSpawn()
    const pending = printPdfViaCups(new Uint8Array([1]), {}, spawnFn as unknown as typeof spawn)
    await vi.waitFor(() => expect(calls).toHaveLength(1))
    calls[0]!.fake.close(3)
    await expect(pending).resolves.toEqual({ ok: false, error: 'lp exited with code 3' })
  })
})

describe('lpPageRange', () => {
  it('collapses consecutive runs and drops noise', () => {
    expect(lpPageRange([1, 2, 3, 7])).toBe('1-3,7')
    expect(lpPageRange([5])).toBe('5')
    expect(lpPageRange([10, 2, 2, 1])).toBe('1-2,10')
    // beyond any document: filtered to nothing rather than printed as garbage
    expect(lpPageRange([0, -3, 1.5])).toBeUndefined()
    expect(lpPageRange([])).toBeUndefined()
  })
})

describe('printFileViaCups', () => {
  it('spools the file operand after the flags, with the -p selection', async () => {
    const { calls, spawnFn } = fakeSpawn()
    const pending = printFileViaCups(
      '/docs/report.pdf',
      { deviceName: 'HP', copies: 2 },
      [4, 1, 2, 3],
      spawnFn as unknown as typeof spawn,
    )
    await vi.waitFor(() => expect(calls).toHaveLength(1))
    expect(calls[0]!.args).toEqual(['-d', 'HP', '-n', '2', '-p', '1-4', '--', '/docs/report.pdf'])
    calls[0]!.fake.close(0)
    await expect(pending).resolves.toEqual({ ok: true })
  })

  it('prints the whole document without a -p flag when no pages are given', async () => {
    const { calls, spawnFn } = fakeSpawn()
    const pending = printFileViaCups(
      '/docs/report.pdf',
      {},
      undefined,
      spawnFn as unknown as typeof spawn,
    )
    await vi.waitFor(() => expect(calls).toHaveLength(1))
    expect(calls[0]!.args).toEqual(['--', '/docs/report.pdf'])
    calls[0]!.fake.stderr.write('lp: printer refused the job\n')
    calls[0]!.fake.close(1)
    await expect(pending).resolves.toEqual({ ok: false, error: 'lp: printer refused the job' })
  })
})

describe('printOutcome', () => {
  it('a cancel in the interactive dialog is not an error, a silent one is', () => {
    expect(printOutcome(false, 'Print job canceled', false)).toEqual({ ok: false })
    expect(printOutcome(false, 'Print job canceled', true)).toEqual({
      ok: false,
      error: 'Print job canceled',
    })
  })

  it('success passes and real failures surface their reason', () => {
    expect(printOutcome(true, undefined, false)).toEqual({ ok: true })
    expect(printOutcome(false, 'Print job failed', false)).toEqual({
      ok: false,
      error: 'Print job failed',
    })
    expect(printOutcome(false, undefined, true)).toEqual({ ok: false })
  })
})
