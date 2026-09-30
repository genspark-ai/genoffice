// Range-resumable installer downloads for the auto-updater (genoffice#…):
// electron-updater writes the installer with a fresh createWriteStream and no
// Range header, so every retry after a dropped connection starts from byte 0 —
// on a slow link a large NSIS/zip that keeps dying at 90% never lands.
//
// This wraps the updater's public httpExecutor.download: the byte-landing part
// becomes "append to <destination>.part with a Range request, verify, rename",
// while everything else (cache registration, signature verification,
// quitAndInstall) keeps running stock electron-updater code. Failures are
// best-effort — anything the wrapper cannot handle falls back to the original
// implementation, so an update can never get *worse* because of this file.
import { createHash } from 'node:crypto'
import { createReadStream, createWriteStream } from 'node:fs'
import { open, rename, stat, unlink } from 'node:fs/promises'
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
/**
 * A wedged connection (gateway black hole) never errors the stream by itself.
 * The stock httpExecutor fails a stalled socket after 60 s; the fetch-based
 * wrapper must match that or a silent hang replaces a retryable failure.
 */
const STALL_TIMEOUT_MS = 60_000

/**
 * Install the resumable download on an electron-updater instance by replacing
 * its httpExecutor.download (a public method on ElectronHttpExecutor). Safe to
 * call once per updater; exported for tests.
 */
export function installResumeDownload(
  updater: { httpExecutor: DownloadExecutor },
  /** shorter stall window for tests; production uses STALL_TIMEOUT_MS */
  stallTimeoutMs: number = STALL_TIMEOUT_MS,
): void {
  const executor = updater.httpExecutor
  // a no-op on any unexpected shape (test doubles, a future electron-updater
  // restructuring): resume is an enhancement, never a load-bearing feature
  if (!executor || typeof executor.download !== 'function') return
  const original = executor.download.bind(executor)
  executor.download = (url, destination, options) =>
    resumeDownload(url, destination, options ?? {}, original, stallTimeoutMs)
}

interface PartMeta {
  /** validator for the bytes already in the .part file: ETag, else Last-Modified */
  ifRange?: string
}

const readPart = async (partPath: string): Promise<{ resumeFrom: number; meta: PartMeta }> => {
  try {
    const [size, meta] = await Promise.all([
      stat(partPath).then((s) => s.size),
      readPartMeta(partPath + '.json'),
    ])
    return { resumeFrom: size > 0 ? size : 0, meta }
  } catch {
    return { resumeFrom: 0, meta: {} }
  }
}

async function readPartMeta(metaPath: string): Promise<PartMeta> {
  try {
    const raw = JSON.parse((await open(metaPath, 'r').then((f) => f.readFile())).toString('utf8'))
    if (raw && typeof raw === 'object' && typeof raw.ifRange === 'string')
      return { ifRange: raw.ifRange }
  } catch {
    /* no usable meta: the next request just goes out without If-Range */
  }
  return {}
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

/**
 * One installer download, resuming `<destination>.part` when the server
 * honours Range. Falls back to the stock implementation when there is no
 * checksum to validate a resumed file against, or when the request itself
 * cannot be placed — a mid-stream failure is a real download failure and is
 * thrown with the .part kept for the next attempt.
 */
async function resumeDownload(
  url: URL,
  destination: string,
  options: ExecutorDownloadOptions,
  fallback: (url: URL, destination: string, options: ExecutorDownloadOptions) => Promise<string>,
  stallTimeoutMs: number,
): Promise<string> {
  // without a checksum a resumed file cannot be validated: keep stock behaviour
  if (!options.sha512) return fallback(url, destination, options)
  const partPath = `${destination}.part`
  const metaPath = `${partPath}.json`
  const encoding = sha512Encoding(options.sha512)

  let resumeFrom = 0
  let ifRange: string | undefined
  try {
    const part = await readPart(partPath)
    resumeFrom = part.resumeFrom
    ifRange = part.meta.ifRange
  } catch {
    /* unreadable part state: download from scratch */
  }

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
      if (ifRange) headers['If-Range'] = ifRange
    }
    cancelCheck()
    response = await fetch(url, {
      headers,
      redirect: 'follow',
      signal: abort.signal,
    })
  } catch (error) {
    if (error instanceof DownloadAbortedError) throw error
    // the request could not even be placed — let the stock stack (electron net,
    // its proxy/session handling) try instead of turning resume into a regression
    return fallback(url, destination, options)
  }

  // a stale .part larger than the artifact: wipe and download whole
  if (response.status === 416 && resumeFrom > 0) {
    await discardPart(partPath, metaPath)
    resumeFrom = 0
    delete headers.Range
    delete headers['If-Range']
    cancelCheck()
    response = await fetch(url, { headers, redirect: 'follow', signal: abort.signal })
  }

  if (!response.ok && response.status !== 206) {
    throw new Error(`Cannot download "${url.href}", status ${response.status}`)
  }

  const appending = response.status === 206 && resumeFrom > 0
  if (!appending) resumeFrom = 0

  // record the response validator for the next attempt before streaming
  const etag = response.headers.get('etag')
  const lastModified = response.headers.get('last-modified')
  const nextIfRange = etag ?? lastModified ?? undefined
  if (nextIfRange) {
    await open(metaPath, 'w').then((f) =>
      f.writeFile(JSON.stringify({ ifRange: nextIfRange })).then(
        () => f.close(),
        () => f.close(),
      ),
    )
  }

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
    if (actual !== options.sha512) {
      await discardPart(partPath, metaPath)
      throw new Error(
        `sha512 checksum mismatch, expected ${options.sha512}, got ${actual} (resumed download discarded)`,
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

async function discardFile(path: string): Promise<void> {
  try {
    await unlink(path)
  } catch {
    /* already gone */
  }
}
