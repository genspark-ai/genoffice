import { mkdtemp, readFile, writeFile, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { ReadableStream } from 'node:stream/web'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { installResumeDownload } from '../src/main/update-resume'

/**
 * The resumable download is exercised against the real filesystem (a temp dir)
 * with fetch mocked at the boundary; the stock electron-updater flow past the
 * download (checksum, signature, install) is untouched and untested here.
 */

const realFetch = globalThis.fetch
let dir: string

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'genoffice-resume-'))
})
afterEach(async () => {
  globalThis.fetch = realFetch
  await rm(dir, { recursive: true, force: true })
})

const sha512 = (data: Buffer | string): string => createHash('sha512').update(data).digest('base64')

const respond = (
  status: number,
  body: Buffer | string,
  headers: Record<string, string> = {},
): Response =>
  new Response(new Uint8Array(Buffer.isBuffer(body) ? body : Buffer.from(body)), {
    status,
    headers,
  })

/** the bytes a fake artifact splits into: prefix (already on disk) + rest */
const PREFIX = Buffer.from('0123456789')
const REST = Buffer.from('abcdefghijklmnopqrstuvwxyz')
const FULL = Buffer.concat([PREFIX, REST])

/** a minimal updater double whose httpExecutor records the fallback calls */
function makeUpdater(stallTimeoutMs?: number) {
  const calls: Array<{ url: string; destination: string }> = []
  const executor = {
    download: async (url: URL, destination: string) => {
      calls.push({ url: url.href, destination })
      return destination
    },
  }
  installResumeDownload(
    { httpExecutor: executor },
    // 50 ms: the stall window is behaviour under test, not the 60 s value
    stallTimeoutMs ?? 50,
  )
  return { executor, fallbackCalls: calls }
}

const url = new URL('https://cdn.example.test/genoffice-setup-1.2.3.exe')
const destOf = (): string => join(dir, 'installer.exe')

describe('resumable installer download', () => {
  it('downloads whole when no .part exists and renames it into place', async () => {
    const { executor } = makeUpdater()
    const fetchMock = vi.fn(async () =>
      respond(200, FULL, { 'content-length': String(FULL.length), etag: '"v1"' }),
    )
    globalThis.fetch = fetchMock as unknown as typeof fetch
    const progress: number[] = []
    const result = await executor.download(url, destOf(), {
      sha512: sha512(FULL),
      onProgress: (p) => progress.push(p.percent),
    })
    expect(result).toBe(destOf())
    expect(await readFile(destOf())).toEqual(FULL)
    // the .part and its meta are gone after the rename
    await expect(stat(destOf() + '.part')).rejects.toThrow()
    await expect(stat(destOf() + '.part.json')).rejects.toThrow()
    expect(progress.at(-1)).toBe(100)
  })

  it('resumes from the .part with Range + If-Range and appends', async () => {
    const dest = destOf()
    await writeFile(dest + '.part', PREFIX)
    await writeFile(dest + '.part.json', JSON.stringify({ ifRange: '"v1"' }))
    const { executor } = makeUpdater()
    const fetchMock = vi.fn(async () =>
      respond(206, REST, { 'content-length': String(REST.length), etag: '"v1"' }),
    )
    globalThis.fetch = fetchMock as unknown as typeof fetch
    const result = await executor.download(url, dest, { sha512: sha512(FULL) })
    expect(result).toBe(dest)
    expect(await readFile(dest)).toEqual(FULL)
    const init = (fetchMock.mock.calls[0]![1] as RequestInit).headers as Record<string, string>
    expect(init.Range).toBe(`bytes=${PREFIX.length}-`)
    expect(init['If-Range']).toBe('"v1"')
  })

  it('a stale If-Range answer (200 whole-file) restarts from byte 0', async () => {
    const dest = destOf()
    await writeFile(dest + '.part', Buffer.from('STALE-GARBAGE'))
    await writeFile(dest + '.part.json', JSON.stringify({ ifRange: '"old"' }))
    const { executor } = makeUpdater()
    const NEW = Buffer.from('a brand new artifact')
    globalThis.fetch = vi.fn(async () => respond(200, NEW)) as unknown as typeof fetch
    await executor.download(url, dest, { sha512: sha512(NEW) })
    expect(await readFile(dest)).toEqual(NEW)
  })

  it('a 416 (part larger than the artifact) wipes the .part and downloads whole', async () => {
    const dest = destOf()
    await writeFile(dest + '.part', Buffer.alloc(FULL.length + 50, 1))
    const { executor } = makeUpdater()
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const headers = (init?.headers ?? {}) as Record<string, string>
      if (headers.Range) return respond(416, Buffer.alloc(0))
      return respond(200, FULL, { 'content-length': String(FULL.length) })
    })
    globalThis.fetch = fetchMock as unknown as typeof fetch
    await executor.download(url, dest, { sha512: sha512(FULL) })
    expect(await readFile(dest)).toEqual(FULL)
    // the retried request carried no Range header
    const second = (fetchMock.mock.calls[1]![1] as RequestInit).headers as Record<string, string>
    expect(second.Range).toBeUndefined()
  })

  it('a checksum mismatch discards the .part and throws', async () => {
    const dest = destOf()
    await writeFile(dest + '.part', PREFIX)
    const { executor } = makeUpdater()
    globalThis.fetch = vi.fn(async () =>
      respond(206, REST, { 'content-length': String(REST.length) }),
    ) as unknown as typeof fetch
    // sha512 of something else entirely
    await expect(executor.download(url, dest, { sha512: sha512('other bytes') })).rejects.toThrow(
      /checksum mismatch/,
    )
    await expect(stat(dest + '.part')).rejects.toThrow()
    await expect(stat(dest)).rejects.toThrow()
  })

  it('a mid-stream network failure keeps the .part for the next attempt', async () => {
    const dest = destOf()
    const { executor } = makeUpdater()
    // a body stream that errors halfway through
    // a real socket hang-up errors the stream after the prefix was delivered
    // (a synchronous error would void the queued chunk — Node web streams)
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(FULL.subarray(0, 4)))
        setTimeout(() => controller.error(new Error('socket hang up')), 5)
      },
    })
    globalThis.fetch = vi.fn(
      async () =>
        new Response(stream, { status: 200, headers: { 'content-length': String(FULL.length) } }),
    ) as unknown as typeof fetch
    await expect(executor.download(url, dest, { sha512: sha512(FULL) })).rejects.toThrow(/hang up/)
    // the received prefix survives: the next attempt resumes from it
    expect((await stat(dest + '.part')).size).toBe(4)
    expect(await readFile(dest + '.part')).toEqual(FULL.subarray(0, 4))
  })

  it('aborts a wedged stream with no data arriving and keeps the .part', async () => {
    const dest = destOf()
    // the injected stall window is 50 ms (see makeUpdater)
    const { executor } = makeUpdater()
    // one chunk arrives, then the connection goes silent (no error, no end)
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(FULL.subarray(0, 5)))
      },
    })
    globalThis.fetch = vi.fn(
      async () =>
        new Response(stream, {
          status: 200,
          headers: { 'content-length': String(FULL.length) },
        }),
    ) as unknown as typeof fetch
    await expect(executor.download(url, dest, { sha512: sha512(FULL) })).rejects.toThrow(
      /stalled: no data for/,
    )
    // the received prefix survives for the next attempt
    expect(await readFile(dest + '.part')).toEqual(FULL.subarray(0, 5))
  })

  it('honours the cancellation token mid-stream and keeps the .part', async () => {
    const dest = destOf()
    const { executor } = makeUpdater()
    let cancel = false
    const options = {
      sha512: sha512(FULL),
      cancellationToken: {
        get cancelled() {
          return cancel
        },
      },
      // flip the token once the first chunk has been written: the writer
      // checks before accepting each next chunk
      onProgress: () => {
        cancel = true
      },
    }
    // Enqueue one chunk per pull(), not both from start(): everything queued
    // synchronously in start() is coalesced by Readable.fromWeb into a single
    // read, so the writer's between-chunks cancellation check would only ever
    // run after the whole body had already landed and the test would prove
    // nothing. One chunk per pull() is what makes the boundary real.
    let sent = 0
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (sent === 0) {
          controller.enqueue(new Uint8Array(FULL.subarray(0, 5)))
          sent = 1
        } else {
          controller.enqueue(new Uint8Array(FULL.subarray(5)))
          controller.close()
        }
      },
    })
    globalThis.fetch = vi.fn(
      async () =>
        new Response(stream, { status: 200, headers: { 'content-length': String(FULL.length) } }),
    ) as unknown as typeof fetch
    await expect(executor.download(url, dest, options)).rejects.toThrow(/aborted|cancelled/)
    // exactly the first chunk: the token was checked and honoured at the boundary
    expect((await stat(dest + '.part')).size).toBe(5)
    expect((await stat(dest + '.part')).size).toBeLessThan(FULL.length)
  })

  it('falls back to the stock download without a sha512 to validate against', async () => {
    const { executor, fallbackCalls } = makeUpdater()
    globalThis.fetch = vi.fn(async () => {
      throw new Error('must not be reached')
    }) as unknown as typeof fetch
    const result = await executor.download(url, destOf(), {})
    expect(fallbackCalls).toHaveLength(1)
    expect(result).toBe(destOf())
  })

  it('falls back when the request cannot be placed at all', async () => {
    const dest = destOf()
    await writeFile(dest + '.part', PREFIX)
    const { executor, fallbackCalls } = makeUpdater()
    globalThis.fetch = vi.fn(async () => {
      throw new TypeError('fetch failed')
    }) as unknown as typeof fetch
    await executor.download(url, dest, { sha512: sha512(FULL) })
    expect(fallbackCalls).toHaveLength(1)
  })

  it('surfaces a server refusal instead of falling back', async () => {
    const { executor, fallbackCalls } = makeUpdater()
    globalThis.fetch = vi.fn(async () => respond(403, Buffer.alloc(0))) as unknown as typeof fetch
    await expect(executor.download(url, destOf(), { sha512: sha512(FULL) })).rejects.toThrow(
      /status 403/,
    )
    expect(fallbackCalls).toHaveLength(0)
  })
})
