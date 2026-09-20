import { randomBytes } from 'node:crypto'
import { lookup } from 'node:dns/promises'
import { createWriteStream, existsSync, mkdirSync, rmSync, statSync } from 'node:fs'
import { request as httpRequest, type IncomingMessage } from 'node:http'
import { request as httpsRequest } from 'node:https'
import { BlockList, isIP } from 'node:net'
import { basename, extname, join } from 'node:path'
import { Transform } from 'node:stream'
import { pipeline } from 'node:stream/promises'

/**
 * Files that travel between a remote MCP client and this server. Uploads land
 * under `<root>/uploads`; outputs a tool wrote anywhere are exposed under an id
 * so the client can download them. Entries expire after `ttlMs` without access.
 */
export interface StoredFile {
  id: string
  name: string
  path: string
}

export const MAX_TRANSFER_BYTES = 256 * 1024 * 1024
const DEFAULT_TTL_MS = 60 * 60 * 1000

export const MIME_TYPES: Record<string, string> = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.pdf': 'application/pdf',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
  '.md': 'text/markdown',
  '.html': 'text/html',
  '.csv': 'text/csv',
  '.json': 'application/json',
  '.txt': 'text/plain',
}

export function mimeOf(path: string): string {
  return MIME_TYPES[extname(path).toLowerCase()] ?? 'application/octet-stream'
}

/** A client-supplied file name reduced to one safe path segment. */
export function safeName(raw: string | undefined, fallback = 'file'): string {
  const base = basename((raw ?? '').trim().replace(/\\/g, '/'))
  const clean = base.replace(/[^\w.\- ()]/g, '_').replace(/^\.+/, '')
  return clean === '' ? fallback : clean
}

export class FileStore {
  private readonly entries = new Map<string, StoredFile & { touched: number }>()
  private readonly idsByPath = new Map<string, string>()
  private readonly sweeper: NodeJS.Timeout

  constructor(
    readonly root: string,
    private readonly ttlMs = DEFAULT_TTL_MS,
  ) {
    mkdirSync(join(root, 'uploads'), { recursive: true })
    this.sweeper = setInterval(() => this.sweep(), Math.min(ttlMs, 5 * 60_000))
    this.sweeper.unref()
  }

  /** Where an upload named `name` goes; the caller writes it, then calls expose(). */
  uploadTarget(name: string): string {
    const dir = join(this.root, 'uploads', newId())
    mkdirSync(dir, { recursive: true })
    return join(dir, safeName(name, 'upload.bin'))
  }

  /** Make a file on disk downloadable; the same path always gets the same id. */
  expose(path: string): StoredFile {
    const known = this.idsByPath.get(path)
    if (known) return this.get(known)!
    const entry = { id: newId(), name: basename(path), path, touched: Date.now() }
    this.entries.set(entry.id, entry)
    this.idsByPath.set(path, entry.id)
    return entry
  }

  get(id: string): StoredFile | undefined {
    const entry = this.entries.get(id)
    if (!entry) return undefined
    if (!existsSync(entry.path)) {
      this.drop(entry.id)
      return undefined
    }
    entry.touched = Date.now()
    return entry
  }

  /** `/files/<id>/<name>` of this server → local path; anything else is undefined. */
  resolveOwnUrl(url: string): string | undefined {
    let pathname: string
    try {
      pathname = new URL(url).pathname
    } catch {
      return undefined
    }
    const m = /^\/files\/([a-f0-9]{16})\/[^/]+$/.exec(pathname)
    return m ? this.get(m[1]!)?.path : undefined
  }

  urlFor(file: StoredFile, baseUrl: string): string {
    return `${baseUrl.replace(/\/$/, '')}/files/${file.id}/${encodeURIComponent(file.name)}`
  }

  sweep(now = Date.now()): void {
    for (const entry of [...this.entries.values()]) {
      if (now - entry.touched > this.ttlMs) this.drop(entry.id)
    }
  }

  dispose(): void {
    clearInterval(this.sweeper)
    rmSync(this.root, { recursive: true, force: true })
  }

  private drop(id: string): void {
    const entry = this.entries.get(id)
    if (!entry) return
    this.entries.delete(id)
    this.idsByPath.delete(entry.path)
    if (entry.path.startsWith(join(this.root, 'uploads'))) {
      rmSync(join(entry.path, '..'), { recursive: true, force: true })
    }
  }
}

function newId(): string {
  return randomBytes(8).toString('hex')
}

export function isHttpUrl(value: unknown): value is string {
  return typeof value === 'string' && /^https?:\/\//i.test(value)
}

/** Cloud instance-metadata endpoints: the one address range a server fetching on a client's behalf must never reach. BlockList canonicalises, so IPv4-mapped and expanded IPv6 spellings match too. */
const BLOCKED = new BlockList()
BLOCKED.addSubnet('169.254.0.0', 16, 'ipv4')
BLOCKED.addSubnet('fe80::', 10, 'ipv6')
BLOCKED.addAddress('fd00:ec2::254', 'ipv6')

interface Pinned {
  address: string
  family: 4 | 6
}

/** Resolve once, check every address, and return the list the connection is pinned to. */
async function pinnedAddresses(url: URL): Promise<Pinned[]> {
  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    throw new Error(`only http(s) URLs can be fetched: ${url.href}`)
  }
  const host = url.hostname.replace(/^\[|\]$/g, '')
  if (host === 'metadata.google.internal') throw new Error(`refusing to fetch ${url.href}`)
  const literal = isIP(host)
  const found = literal
    ? [{ address: host, family: literal }]
    : await lookup(host, { all: true, verbatim: true })
  if (found.length === 0) throw new Error(`cannot resolve ${host}`)
  for (const a of found) {
    if (BLOCKED.check(a.address, a.family === 6 ? 'ipv6' : 'ipv4')) {
      throw new Error(`refusing to fetch ${url.href}`)
    }
  }
  return found.map((a) => ({ address: a.address, family: a.family === 6 ? 6 : 4 }))
}

type LookupCallback = (err: null, address: string | Pinned[], family?: number) => void

/** The socket goes to the addresses that passed the check, whatever DNS says now; `all` is what autoSelectFamily asks for. */
function pinnedLookup(pinned: Pinned[]) {
  return (_host: string, opts: { all?: boolean }, cb: LookupCallback): void => {
    if (opts.all) cb(null, pinned)
    else cb(null, pinned[0]!.address, pinned[0]!.family)
  }
}

function get(url: URL, pinned: Pinned[], signal: AbortSignal): Promise<IncomingMessage> {
  return new Promise((resolve, reject) => {
    const request = url.protocol === 'https:' ? httpsRequest : httpRequest
    const req = request(
      url,
      { method: 'GET', signal, lookup: pinnedLookup(pinned) as never },
      resolve,
    )
    req.on('error', reject)
    req.end()
  })
}

export interface FetchOptions {
  timeoutMs?: number
  maxBytes?: number
}

/**
 * Download `url` into a fresh directory under `dir`, keeping the remote file name,
 * and return the local path. Redirects are followed by hand so every hop is
 * resolved, checked and pinned in turn.
 */
export async function fetchToFile(
  url: string,
  dir: string,
  opts: FetchOptions = {},
): Promise<string> {
  const maxBytes = opts.maxBytes ?? MAX_TRANSFER_BYTES
  const signal = AbortSignal.timeout(opts.timeoutMs ?? 60_000)
  let current = new URL(url)
  let response: IncomingMessage | undefined
  for (let hop = 0; hop < 5; hop++) {
    response = await get(current, await pinnedAddresses(current), signal)
    const status = response.statusCode ?? 0
    const location = response.headers.location
    if (status >= 300 && status < 400 && location) {
      response.resume()
      const next = new URL(location, current)
      // Keep redirects http(s)-only: pinnedAddresses would reject them on the
      // next hop, but fail fast here with a clear error instead.
      if (next.protocol !== 'http:' && next.protocol !== 'https:') {
        throw new Error(`only http(s) URLs can be fetched: ${next.href}`)
      }
      current = next
      response = undefined
      continue
    }
    break
  }
  if (!response) throw new Error(`too many redirects fetching ${url}`)
  const status = response.statusCode ?? 0
  if (status < 200 || status >= 300) {
    response.resume()
    throw new Error(`fetching ${url} failed: HTTP ${status}`)
  }
  const declared = Number(response.headers['content-length'] ?? 0)
  if (declared > maxBytes) {
    response.destroy()
    throw new Error(`${url} is larger than ${maxBytes} bytes`)
  }
  const name = remoteName(response, current)
  const target = join(dir, `in-${newId()}`)
  mkdirSync(target, { recursive: true })
  const path = join(target, name)
  let received = 0
  const counted = new Transform({
    transform(chunk: Buffer, _enc, cb) {
      received += chunk.byteLength
      if (received > maxBytes) cb(new Error(`${url} exceeds ${maxBytes} bytes`))
      else cb(null, chunk)
    },
  })
  try {
    await pipeline(response, counted, createWriteStream(path), { signal })
  } catch (err) {
    rmSync(target, { recursive: true, force: true })
    throw err
  }
  if (!statSync(path).isFile()) throw new Error(`fetching ${url} wrote nothing`)
  return path
}

function decodeSafe(value: string): string {
  try {
    return decodeURIComponent(value)
  } catch {
    // Malformed percent-encoding such as %ZZ: keep the raw text and let
    // safeName sanitize it to a safe segment instead of throwing.
    return value
  }
}

function remoteName(response: IncomingMessage, url: URL): string {
  const disposition = response.headers['content-disposition'] ?? ''
  const star = /filename\*=(?:UTF-8'')?([^;]+)/i.exec(disposition)
  const plain = /filename="?([^";]+)"?/i.exec(disposition)
  const fromHeader = star ? decodeSafe(star[1]!.trim()) : plain?.[1]
  const fromPath = decodeSafe(basename(url.pathname))
  let name = safeName(fromHeader ?? fromPath, 'download')
  if (extname(name) === '') {
    const type = (response.headers['content-type'] ?? '').split(';')[0]!.trim()
    const ext = Object.entries(MIME_TYPES).find(([, mime]) => mime === type)?.[0]
    if (ext) name += ext
  }
  return name
}
