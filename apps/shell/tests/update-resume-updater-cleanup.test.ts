import { mkdir, mkdtemp, readdir, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, relative, sep } from 'node:path'
import { createHash } from 'node:crypto'
import { ReadableStream } from 'node:stream/web'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { AppUpdater, NoOpLogger } from 'electron-updater'
import { installResumeDownload } from '../src/main/update-resume'

/**
 * The sibling suite mocks fetch and drives the wrapper directly, which cannot
 * see the defect this file is about: electron-updater downloads into
 * `<cacheDir>/pending/` and then *empties that directory* on the very failure
 * the resume feature exists for.
 *
 * So this suite runs the real thing. `AppUpdater.prototype.executeDownload`
 * from electron-updater 6.8.9 is invoked on a real instance (built from its
 * prototype, with only the collaborators it actually reads faked), and that
 * method creates its own real `DownloadedUpdateHelper`. Every filesystem effect
 * asserted below — the temp file, the `emptyDir(pending/)` in the failure
 * handler, the rename on success — is the dependency's own code. Only `fetch`
 * is faked, because there is no server to talk to.
 */
const realFetch = globalThis.fetch
let dir: string
/** the updater cache dir executeDownload computes: join(baseCachePath, updaterCacheDirName) */
let pendingDir: string
let resumeDir: string

const sha512 = (data: Buffer | string): string => createHash('sha512').update(data).digest('base64')

/**
 * The prefix has to be larger than the reader's 16 KiB high-water mark, and the
 * stream has to be a byte stream. `Readable.fromWeb` pulls ahead to fill its
 * buffer, so a small first chunk is followed by the next `pull` — and a web
 * stream discards its queued chunks when it errors — before the writer ever
 * sees anything. A byte stream has a high-water mark of 0, so `pull` only runs
 * again once the previous chunk has been read: the prefix is written, then the
 * connection dies. 64 KiB clears the Node-side mark on the versions this runs
 * on. Same forcing trick as the cancellation test, inverted.
 */
const HEAD = 64 * 1024 + 7
const PREFIX = Buffer.alloc(HEAD, 0x61)
const TAIL = Buffer.from('tail')
const FULL = Buffer.concat([PREFIX, TAIL])

/** a body that delivers the prefix and then dies, deterministically */
function dyingBody(prefix: Buffer): ReadableStream<Uint8Array> {
  let pulled = false
  return new ReadableStream<Uint8Array>({
    type: 'bytes',
    pull(controller) {
      if (pulled) controller.error(new Error('socket hang up'))
      else {
        pulled = true
        controller.enqueue(new Uint8Array(prefix))
      }
    },
  })
}

/** the artifact url; its trailing extension must match `fileExtension` (see getCacheUpdateFileName) */
const artifactUrl = new URL('https://cdn.example.test/genoffice-mac-1.2.3.zip')
/** what electron-updater names the file it downloads, and renames the temp file to */
const artifactName = 'genoffice-mac-1.2.3.zip'
const tempName = `temp-${artifactName}`

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'genoffice-resume-e2e-'))
  const cacheDir = join(dir, 'updater-cache')
  pendingDir = join(cacheDir, 'pending')
  resumeDir = join(cacheDir, 'resume')
})
afterEach(async () => {
  globalThis.fetch = realFetch
  await rm(dir, { recursive: true, force: true })
})

/**
 * A real AppUpdater with only what executeDownload reads replaced: the logger,
 * the app path/identity that locates the cache dir, the on-disk config that
 * names the cache subdir, and the http executor. `executeDownload` creates the
 * real DownloadedUpdateHelper itself, so the cleanup under test belongs to the
 * dependency rather than to a stand-in.
 */
function makeRealUpdater(): { updater: AppUpdater; stockDownloads: string[] } {
  const stockDownloads: string[] = []
  const stockExecutor = {
    download: async (_url: URL, destination: string) => {
      stockDownloads.push(destination)
      return destination
    },
  }
  const updater = Object.create(AppUpdater.prototype) as AppUpdater
  const bag = updater as unknown as Record<string, unknown>
  bag._logger = new NoOpLogger()
  bag._events = {}
  bag.app = { baseCachePath: dir, name: 'genoffice' }
  bag.configOnDisk = { value: Promise.resolve({ updaterCacheDirName: 'updater-cache' }) }
  bag.httpExecutor = stockExecutor
  // 50 ms stall window so a wedged stream fails fast; the response window is
  // left at its production value because every fetch here answers immediately
  installResumeDownload(updater as unknown as { httpExecutor: typeof stockExecutor }, 50)
  return { updater, stockDownloads }
}

/** the `task` NsisUpdater/MacUpdater hand to executeDownload (their download branch) */
function makeTaskOptions(
  updater: AppUpdater,
  sha: string,
): Parameters<AppUpdater['executeDownload']>[0] {
  const executor = updater as unknown as {
    httpExecutor: { download: (u: URL, d: string, o: unknown) => Promise<string> }
  }
  return {
    fileExtension: 'zip',
    fileInfo: { url: artifactUrl, info: { sha512: sha }, packageInfo: null },
    downloadUpdateOptions: {
      requestHeaders: {},
      cancellationToken: { cancelled: false },
      updateInfoAndProvider: { info: { version: '1.2.3', path: artifactName, sha512: sha } },
    },
    task: async (destinationFile, downloadOptions) =>
      executor.httpExecutor.download(artifactUrl, destinationFile, downloadOptions),
    done: async () => [join(pendingDir, artifactName)],
  } as unknown as Parameters<AppUpdater['executeDownload']>[0]
}

/** the stored resume state, as the filesystem actually holds it */
async function storedParts(): Promise<string[]> {
  try {
    return (await readdir(resumeDir)).sort()
  } catch {
    return []
  }
}

describe('resume state vs. the updater emptying its own cache dir', () => {
  it('keeps the .part when executeDownload fails and empties pending/, then resumes on the retry', async () => {
    const { updater, stockDownloads } = makeRealUpdater()

    // attempt 1: the connection dies after a prefix has reached the disk
    globalThis.fetch = vi.fn(
      async () =>
        new Response(dyingBody(PREFIX), {
          status: 200,
          headers: { 'content-length': String(FULL.length), etag: '"v1"' },
        }),
    ) as unknown as typeof fetch

    await expect(updater.executeDownload(makeTaskOptions(updater, sha512(FULL)))).rejects.toThrow(
      /hang up/,
    )

    // the control assertion: the real failure handler really did empty pending/
    // (removeFileIfAny -> clear() -> emptyDir(cacheDirForPendingUpdate))
    expect(await readdir(pendingDir)).toEqual([])
    // and the stock executor was never reached: this was a real download failure
    expect(stockDownloads).toEqual([])

    // the bytes that did arrive are still on disk, outside pending/
    const partPath = join(resumeDir, `${tempName}.part`)
    expect(await storedParts()).toEqual([`${tempName}.part`, `${tempName}.part.json`])
    expect((await stat(partPath)).size).toBe(HEAD)
    expect(await readFile(partPath)).toEqual(PREFIX)
    // nothing of ours is left where the updater deletes things
    expect(relative(pendingDir, partPath).startsWith('..')).toBe(true)
    expect(relative(resumeDir, partPath).split(sep)).toEqual([`${tempName}.part`])

    // attempt 2: the server honours Range and the update lands
    const fetchMock = vi.fn(
      async () =>
        new Response(new Uint8Array(TAIL), {
          status: 206,
          headers: { 'content-length': String(TAIL.length), etag: '"v1"' },
        }),
    )
    globalThis.fetch = fetchMock as unknown as typeof fetch

    await updater.executeDownload(makeTaskOptions(updater, sha512(FULL)))

    const headers = (fetchMock.mock.calls[0]![1] as RequestInit).headers as Record<string, string>
    expect(headers.Range).toBe(`bytes=${HEAD}-`)
    expect(headers['If-Range']).toBe('"v1"')
    // the completed artifact is registered under pending/ where the updater expects it
    expect(await readFile(join(pendingDir, artifactName))).toEqual(FULL)
    // the resume directory is left clean
    expect(await storedParts()).toEqual([])
  })

  it('refuses to resume a .part that belongs to another artifact', async () => {
    const { updater } = makeRealUpdater()

    // attempt 1 stores a partial file, pinned to the digest it belongs to
    globalThis.fetch = vi.fn(
      async () =>
        new Response(dyingBody(PREFIX), {
          status: 200,
          headers: { 'content-length': String(FULL.length) },
        }),
    ) as unknown as typeof fetch
    await expect(updater.executeDownload(makeTaskOptions(updater, sha512(FULL)))).rejects.toThrow(
      /hang up/,
    )
    expect((await stat(join(resumeDir, `${tempName}.part`))).size).toBe(HEAD)

    // attempt 2 is for a different artifact at the same path: the stored bytes
    // are not proven to belong to it, so no Range header may be sent
    const NEW = Buffer.from('a brand new artifact')
    const fetchMock = vi.fn(
      async () =>
        new Response(new Uint8Array(NEW), {
          status: 200,
          headers: { 'content-length': String(NEW.length) },
        }),
    )
    globalThis.fetch = fetchMock as unknown as typeof fetch
    await updater.executeDownload(makeTaskOptions(updater, sha512(NEW)))
    const headers = (fetchMock.mock.calls[0]![1] as RequestInit).headers as Record<string, string>
    expect(headers.Range).toBeUndefined()
    expect(await readFile(join(pendingDir, artifactName))).toEqual(NEW)
  })

  it('drops resume state belonging to a superseded artifact so it cannot pile up', async () => {
    const { updater } = makeRealUpdater()
    // a part left behind by an earlier update that never completed
    await mkdir(resumeDir, { recursive: true })
    const stale = 'temp-genoffice-mac-1.2.2.zip'
    await writeFile(join(resumeDir, `${stale}.part`), PREFIX)
    await writeFile(
      join(resumeDir, `${stale}.part.json`),
      JSON.stringify({ sha512: sha512(PREFIX) }),
    )

    globalThis.fetch = vi.fn(
      async () =>
        new Response(new Uint8Array(FULL), {
          status: 200,
          headers: { 'content-length': String(FULL.length) },
        }),
    ) as unknown as typeof fetch
    await updater.executeDownload(makeTaskOptions(updater, sha512(FULL)))
    expect(await storedParts()).toEqual([])
  })
})
