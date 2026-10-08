// Range-resumable installer downloads (genoffice#1777): electron-updater
// restarts every retry from byte 0. This wraps httpExecutor.download so the
// bytes land in <resume>/<artifact>.part via Range requests; anything the
// wrapper cannot handle falls back to the stock implementation.
import { createHash } from 'node:crypto'
import { createReadStream, createWriteStream } from 'node:fs'
import { mkdir, readdir, readFile, rename, stat, unlink, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { Readable } from 'node:stream'

/** the slices of builder-util-runtime's DownloadOptions / CancellationToken the wrapper reads */
export interface ExecutorDownloadOptions {
  headers?: Record<string, string | null | undefined> | null
  sha512?: string | null
  cancellationToken?: { readonly cancelled: boolean }
  onProgress?: (info: {
    total: number
    delta: number
    transferred: number
    percent: number
  }) => void
}

export interface DownloadExecutor {
  download(url: URL, destination: string, options: ExecutorDownloadOptions): Promise<string>
}

/** progress events are noisier than the stock transform; keep the UI cadence */
const PROGRESS_INTERVAL_MS = 250
// matches the stock httpExecutor's stalled-socket timeout
const STALL_TIMEOUT_MS = 60_000
// undici's fetch has no timeout and the stall watchdog only arms once a response exists
const RESPONSE_TIMEOUT_MS = 30_000

export function installResumeDownload(
  updater: { httpExecutor: DownloadExecutor },
  /** shorter stall window for tests; production uses STALL_TIMEOUT_MS */
  stallTimeoutMs: number = STALL_TIMEOUT_MS,
  /** shorter response window for tests; production uses RESPONSE_TIMEOUT_MS */
  responseTimeoutMs: number = RESPONSE_TIMEOUT_MS,
): void {
  const executor = updater.httpExecutor
  // a no-op on any unexpected shape (test doubles, a future electron-updater
  // restructuring): resume is an enhancement, never a load-bearing feature
  if (!executor || typeof executor.download !== 'function') return
  const original = executor.download.bind(executor)
  executor.download = (url, destination, options) =>
    resumeDownload(url, destination, options ?? {}, original, stallTimeoutMs, responseTimeoutMs)
}

// A sibling of `pending/`, not inside it: electron-updater empties
// `cacheDirForPendingUpdate` on every download failure and sha mismatch, which
// is exactly the failure the .part must survive. The parent cacheDir is never
// emptied and is per-app and OS-purgeable.
const RESUME_DIR_NAME = 'resume'

interface PartPaths {
  /** the `resume/` directory holding the part and its meta */
  dir: string
  partPath: string
  metaPath: string
  /** the part's base name, used to keep it out of the prune sweep */
  base: string
}

const resumePaths = (destination: string): PartPaths => {
  const destinationDir = path.dirname(destination)
  const dir = path.join(path.dirname(destinationDir), RESUME_DIR_NAME)
  const base = path.basename(destination)
  return {
    dir,
    partPath: path.join(dir, `${base}.part`),
    metaPath: path.join(dir, `${base}.part.json`),
    base,
  }
}

interface PartMeta {
  /** validator for the bytes already in the .part file: ETag, else Last-Modified */
  ifRange?: string
  /** pins the part to one artifact so a redeployed build is never appended to */
  sha512?: string
}

// anything unproven restarts from byte 0: a wrong guess would stitch two versions together
const readPart = async (
  partPath: string,
  metaPath: string,
  sha512: string,
): Promise<{ resumeFrom: number; meta: PartMeta }> => {
  const empty = { resumeFrom: 0, meta: {} as PartMeta }
  try {
    const meta = await readPartMeta(metaPath)
    if (meta.sha512 !== sha512) {
      // stale part for a different artifact: drop it rather than keep it around
      await discardPart(partPath, metaPath)
      return empty
    }
    const size = await stat(partPath).then((s) => s.size)
    return { resumeFrom: size > 0 ? size : 0, meta }
  } catch {
    return empty
  }
}

async function readPartMeta(metaPath: string): Promise<PartMeta> {
  try {
    // readFile(path), not FileHandle.readFile(): the latter leaves the handle
    // open, and an unclosed FileHandle that reaches GC is a hard error on Node
    const raw = JSON.parse(await readFile(metaPath, 'utf8'))
    if (raw && typeof raw === 'object') {
      // each field stands on its own: a server that sends no ETag or
      // Last-Modified still yields a usable (digest-pinned) part, it just
      // cannot offer an If-Range validator
      const meta: PartMeta = {}
      if (typeof raw.sha512 === 'string') meta.sha512 = raw.sha512
      if (typeof raw.ifRange === 'string') meta.ifRange = raw.ifRange
      return meta
    }
  } catch {
    /* no usable meta: the next request just goes out without If-Range */
  }
  return {}
}

async function writePartMeta(metaPath: string, meta: PartMeta): Promise<void> {
  await writeFile(metaPath, JSON.stringify(meta))
}

// parts of superseded versions would otherwise accumulate for the life of the install
async function pruneOtherParts(dir: string, keepBase: string): Promise<void> {
  try {
    const keep = new Set([`${keepBase}.part`, `${keepBase}.part.json`])
    const entries = await readdir(dir)
    await Promise.all(
      entries.filter((name) => !keep.has(name)).map((name) => discardFile(path.join(dir, name))),
    )
  } catch {
    /* no resume dir yet, or unreadable: nothing to prune */
  }
}

/** sha512 encoding exactly as builder-util-runtime's DigestTransform picks it */
function sha512Encoding(sha512: string): 'hex' | 'base64' {
  return sha512.length === 128 &&
    !sha512.includes('+') &&
    !sha512.includes('Z') &&
    !sha512.includes('=')
    ? 'hex'
    : 'base64'
}

async function sha512Of(file: string, encoding: 'hex' | 'base64'): Promise<string> {
  const hash = createHash('sha512')
  for await (const chunk of createReadStream(file)) hash.update(chunk as Buffer)
  return hash.digest(encoding)
}

class DownloadAbortedError extends Error {
  constructor(reason: string) {
    super(`resumable download aborted: ${reason}`)
  }
}

/** the request phase failed for a reason that is not about the artifact */
class RequestFailedError extends Error {
  constructor(reason: string) {
    super(`resumable download could not place its request: ${reason}`)
  }
}

async function placeRequest(
  url: URL,
  headers: Record<string, string>,
  abort: AbortController,
  timeoutMs: number,
): Promise<Response> {
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    return await Promise.race([
      fetch(url, { headers, redirect: 'follow', signal: abort.signal }),
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => {
          abort.abort()
          reject(new RequestFailedError(`no response within ${timeoutMs / 1000} s`))
        }, timeoutMs)
      }),
    ])
  } finally {
    if (timer) clearTimeout(timer)
  }
}

// A mid-stream failure is thrown with the .part kept; anything the wrapper
// cannot own (no checksum, unplaceable request, odd status) falls back to stock.
async function resumeDownload(
  url: URL,
  destination: string,
  options: ExecutorDownloadOptions,
  fallback: (url: URL, destination: string, options: ExecutorDownloadOptions) => Promise<string>,
  stallTimeoutMs: number,
  responseTimeoutMs: number,
): Promise<string> {
  // without a checksum a resumed file cannot be validated: keep stock behaviour
  if (!options.sha512) return fallback(url, destination, options)
  const expectedSha512 = options.sha512
  const { dir, partPath, metaPath, base } = resumePaths(destination)
  const encoding = sha512Encoding(expectedSha512)

  await mkdir(dir, { recursive: true })
  await pruneOtherParts(dir, base)

  const part = await readPart(partPath, metaPath, expectedSha512)
  // reassigned below: the 416 retry and the non-appending case both reset it
  let resumeFrom = part.resumeFrom
  const ifRange: string | undefined = part.meta.ifRange

  const headers: Record<string, string> = {}
  for (const [key, value] of Object.entries(options.headers ?? {}))
    if (typeof value === 'string') headers[key] = value

  const abort = new AbortController()
  const cancelCheck = (): void => {
    if (options.cancellationToken?.cancelled) {
      abort.abort()
      throw new DownloadAbortedError('cancelled')
    }
  }

  let response: Response
  try {
    if (resumeFrom > 0) {
      headers.Range = `bytes=${resumeFrom}-`
      // If-Range guarantees the appended bytes belong to the file whose prefix
      // is on disk; without it a redeployed artifact would stitch two versions
      // (the .part's own sha512 pin already refuses that case outright)
      if (ifRange) headers['If-Range'] = ifRange
    }
    cancelCheck()
    response = await placeRequest(url, headers, abort, responseTimeoutMs)
  } catch (error) {
    // a cancellation during the wait is a cancellation, not a placement failure
    cancelCheck()
    if (error instanceof DownloadAbortedError) throw error
    // the request could not be placed at all (proxy black hole, DNS, TLS,
    // offline) — let the stock stack (electron net, its proxy/session handling,
    // its own timeouts) try instead of turning resume into a regression
    return fallback(url, destination, options)
  }

  // a stale .part larger than the artifact: wipe and download whole
  if (response.status === 416 && resumeFrom > 0) {
    await discardPart(partPath, metaPath)
    resumeFrom = 0
    delete headers.Range
    delete headers['If-Range']
    try {
      cancelCheck()
      response = await placeRequest(url, headers, abort, responseTimeoutMs)
    } catch (error) {
      cancelCheck()
      if (error instanceof DownloadAbortedError) throw error
      return fallback(url, destination, options)
    }
  }

  // undici ignores the Electron session proxy, so a proxy 407 / captive-portal
  // 403 lands here; the stock executor (electron net) handles those
  if (!response.ok && response.status !== 206) {
    return fallback(url, destination, options)
  }

  const appending = response.status === 206 && resumeFrom > 0
  if (!appending) resumeFrom = 0

  // record the artifact digest (always) and the response validator for the next
  // attempt (when the server offers one) before any byte is streamed
  const etag = response.headers.get('etag')
  const lastModified = response.headers.get('last-modified')
  const nextIfRange: string | undefined = etag ?? lastModified ?? undefined
  await writePartMeta(metaPath, {
    sha512: expectedSha512,
    ...(nextIfRange ? { ifRange: nextIfRange } : {}),
  })

  const contentLength = Number(response.headers.get('content-length') ?? '0')
  const total = resumeFrom + (Number.isFinite(contentLength) ? contentLength : 0)
  let transferred = resumeFrom
  let lastProgressAt = 0
  const reportProgress = (delta: number, force = false): void => {
    if (!options.onProgress) return
    const now = Date.now()
    if (!force && now - lastProgressAt < PROGRESS_INTERVAL_MS) return
    lastProgressAt = now
    options.onProgress({
      total,
      delta,
      transferred,
      percent: total > 0 ? (transferred / total) * 100 : 0,
    })
  }

  const fileOut = createWriteStream(partPath, appending ? { flags: 'r+', start: resumeFrom } : {})
  /** wait for buffered writes to reach the disk; on failure keep what made it */
  const settle = (): Promise<void> => new Promise((resolve) => fileOut.end(() => resolve()))
  // re-armed on every chunk (clear+set, not refresh, so fake timers work too)
  let stallTimer: ReturnType<typeof setTimeout> | undefined
  let stalled = false
  let streamIn: Readable | undefined
  const onStall = (): void => {
    stalled = true
    abort.abort()
    // a hand-built Response body ignores the signal; destroy the stream directly
    streamIn?.destroy(new DownloadAbortedError('stalled'))
  }
  const armStallWatchdog = (): void => {
    if (stallTimer) clearTimeout(stallTimer)
    stallTimer = setTimeout(onStall, stallTimeoutMs)
  }
  try {
    if (!response.body) throw new DownloadAbortedError('empty body')
    streamIn = Readable.fromWeb(response.body as Parameters<typeof Readable.fromWeb>[0])
    // a manual loop (not pipeline): the write stream must never be destroyed
    // with writes still buffered, or a dropped connection loses the very
    // bytes this file exists to keep. Backpressure via drain.
    armStallWatchdog()
    for await (const chunk of streamIn) {
      cancelCheck()
      armStallWatchdog()
      const buf = chunk as Buffer
      if (!fileOut.write(buf)) {
        await new Promise<void>((resolve) => fileOut.once('drain', resolve))
      }
      transferred += buf.length
      reportProgress(buf.length)
    }
    if (stallTimer) clearTimeout(stallTimer)
    cancelCheck()
    await settle()
    reportProgress(0, true)

    const actual = await sha512Of(partPath, encoding)
    if (actual !== expectedSha512) {
      await discardPart(partPath, metaPath)
      throw new Error(
        `sha512 checksum mismatch, expected ${expectedSha512}, got ${actual} (resumed download discarded)`,
      )
    }
    await discardFile(metaPath).catch(() => {})
    await rename(partPath, destination)
    return destination
  } catch (error) {
    // the .part stays (what has reached the disk): the next attempt resumes
    if (stallTimer) clearTimeout(stallTimer)
    abort.abort()
    await settle().catch(() => {})
    if (stalled)
      throw new DownloadAbortedError(`download stalled: no data for ${stallTimeoutMs / 1000} s`)
    throw error instanceof Error ? error : new Error(String(error))
  }
}

async function discardPart(partPath: string, metaPath: string): Promise<void> {
  await Promise.all([discardFile(partPath), discardFile(metaPath)])
}

async function discardFile(filePath: string): Promise<void> {
  try {
    await unlink(filePath)
  } catch {
    /* already gone */
  }
}
