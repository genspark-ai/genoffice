import { createHash } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { open, stat } from 'node:fs/promises'
import { inflateRawSync } from 'node:zlib'
import { protocol } from 'electron'
import { DOCX_MEDIA_SCHEME_PRIVILEGE } from '@genoffice/electron-utils'
import {
  LAZY_MEDIA_SCHEME,
  isLazyMediaPart,
  parseLazyMediaUrl,
} from '@genoffice/docx-engine/lazy-media'
import {
  bufferSource,
  lazyMediaHashesIn,
  materializeDocx,
  openZipFile,
  readZipEntries,
  slimDocx,
  type ZipEntry,
  type ZipFile,
} from '@genoffice/docx-engine/zip-splice'

/** documents carrying at least this much browser-decodable media open lazily */
const LAZY_MEDIA_MIN_BYTES = 16 * 1024 * 1024

interface Source {
  path: string
  owners: Set<number>
  table?: { size: number; mtimeMs: number; entries: Map<string, ZipEntry> }
  /** the plain document once the file on disk is encrypted and cannot be served */
  bytes?: Buffer
  /** entry table parsed from exactly `bytes` (identity-checked, never the file's) */
  memTable?: { bytes: Buffer; entries: Map<string, ZipEntry> }
}

/** original-file sha256 → where its media lives */
const sources = new Map<string, Source>()

const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
}

function sha256File(path: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const hash = createHash('sha256')
    createReadStream(path)
      .on('error', reject)
      .on('data', (chunk) => hash.update(chunk))
      .on('end', () => resolve(hash.digest('hex')))
  })
}

function registerSource(hash: string, path: string, wcId: number): void {
  const source = sources.get(hash)
  if (source) {
    if (source.path !== path) source.table = undefined
    source.path = path
    source.owners.add(wcId)
  } else {
    sources.set(hash, { path, owners: new Set([wcId]) })
  }
}

/**
 * The document with its pictures replaced by placeholders, or null when the
 * file should be read whole (small media, encrypted container, odd zip).
 */
export async function openLazyDocx(
  filePath: string,
  wcId: number,
): Promise<{ bytes: Buffer; hash: string } | null> {
  let zip: ZipFile
  try {
    zip = await openZipFile(filePath)
  } catch {
    return null
  }
  try {
    let media = 0
    for (const e of zip.entries.values()) if (isLazyMediaPart(e.name)) media += e.usize
    if (media < LAZY_MEDIA_MIN_BYTES) return null
    const hash = await sha256File(filePath)
    const slim = await slimDocx(zip, hash, LAZY_MEDIA_MIN_BYTES)
    if (!slim) return null
    registerSource(hash, filePath, wcId)
    return { bytes: slim.bytes, hash }
  } catch {
    return null
  } finally {
    await zip.close()
  }
}

/** a recovery copy names the hash its placeholders were cut under; the
 *  on-disk file still holds those parts by name */
export async function adoptLazyMediaHashes(
  bytes: Buffer,
  filePath: string,
  wcId: number,
): Promise<void> {
  try {
    for (const hash of await lazyMediaHashesIn(bytes)) registerSource(hash, filePath, wcId)
  } catch {
    /* not a plain zip: nothing lazy in it */
  }
}

export function forgetLazyMediaOwner(wcId: number): void {
  for (const [hash, source] of sources) {
    source.owners.delete(wcId)
    if (source.owners.size === 0) sources.delete(hash)
  }
}

async function entriesOf(source: Source): Promise<Map<string, ZipEntry> | null> {
  const st = await stat(source.path).catch(() => null)
  if (!st) return null
  if (source.table && source.table.size === st.size && source.table.mtimeMs === st.mtimeMs) {
    return source.table.entries
  }
  const zip = await openZipFile(source.path)
  try {
    // an encrypted save may have switched the source to memory meanwhile
    if (!source.bytes) source.table = { size: st.size, mtimeMs: st.mtimeMs, entries: zip.entries }
    return zip.entries
  } finally {
    await zip.close()
  }
}

async function readRange(path: string, offset: number, length: number): Promise<Buffer> {
  const fh = await open(path, 'r')
  try {
    const buf = Buffer.alloc(length)
    let done = 0
    while (done < length) {
      const { bytesRead } = await fh.read(buf, done, length - done, offset + done)
      if (bytesRead === 0) throw new Error('lazy media: source file truncated')
      done += bytesRead
    }
    return buf
  } finally {
    await fh.close()
  }
}

async function sourceZip(hash: string): Promise<ZipFile | null> {
  const source = sources.get(hash)
  if (source?.bytes) {
    const bytes = source.bytes
    const src = bufferSource(bytes)
    let mem = source.memTable
    if (!mem || mem.bytes !== bytes) {
      const entries = new Map<string, ZipEntry>()
      for (const e of await readZipEntries(src)) entries.set(e.name, e)
      source.memTable = mem = { bytes, entries }
    }
    return { ...src, entries: mem.entries, close: async () => {} }
  }
  const entries = source ? await entriesOf(source) : null
  if (!source || !entries) return null
  return {
    size: (await stat(source.path)).size,
    entries,
    read: (offset, length) => {
      // the file's offsets mean nothing once the source moved to memory
      if (source.bytes) throw new Error('lazy media source replaced during read')
      return readRange(source.path, offset, length)
    },
    close: async () => {},
  }
}

/** saved bytes with lazy placeholders swapped back for the original media,
 *  plus the hashes those placeholders named */
export async function materializeLazyDocx(
  bytes: Buffer,
): Promise<{ bytes: Buffer; hashes: Set<string> }> {
  const hashes = await lazyMediaHashesIn(bytes).catch(() => new Set<string>())
  if (hashes.size === 0) return { bytes, hashes }
  const opened = new Map<string, Promise<ZipFile | null>>()
  const full = await materializeDocx(bytes, (hash) => {
    let zip = opened.get(hash)
    if (!zip) opened.set(hash, (zip = sourceZip(hash)))
    return zip
  })
  return { bytes: full, hashes }
}

/** after a save landed a full copy at `filePath`, serve those hashes from it;
 *  an encrypted save passes the plain bytes instead, since the file cannot be read by range */
export function pointLazyMediaAt(
  hashes: Set<string>,
  filePath: string,
  wcId: number,
  plain?: Buffer,
): void {
  for (const hash of hashes) {
    registerSource(hash, filePath, wcId)
    const source = sources.get(hash)
    if (source) {
      source.bytes = plain
      source.table = undefined
      source.memTable = undefined
    }
  }
}

/** the picture behind a lazy URL, or null when it is not served (any more) */
export async function readLazyMedia(url: string): Promise<{ body: Buffer; mime: string } | null> {
  const ref = parseLazyMediaUrl(url)
  if (!ref || !isLazyMediaPart(ref.partPath)) return null
  const zip = await sourceZip(ref.hash)
  const entry = zip?.entries.get(ref.partPath)
  if (!zip || !entry || (entry.method !== 0 && entry.method !== 8)) return null
  const raw = await zip.read(entry.dataOffset, entry.csize)
  const ext = ref.partPath.split('.').pop()?.toLowerCase() ?? ''
  return {
    body: entry.method === 0 ? raw : inflateRawSync(raw),
    mime: MIME[ext] ?? 'application/octet-stream',
  }
}

export function registerLazyMediaProtocol(): void {
  if (DOCX_MEDIA_SCHEME_PRIVILEGE.scheme !== LAZY_MEDIA_SCHEME) {
    throw new Error('lazy media scheme privilege registered under another name')
  }
  if (protocol.isProtocolHandled(LAZY_MEDIA_SCHEME)) return
  protocol.handle(LAZY_MEDIA_SCHEME, async (request) => {
    const media = await readLazyMedia(request.url)
    if (!media) return new Response(null, { status: 404 })
    return new Response(new Uint8Array(media.body), { headers: { 'Content-Type': media.mime } })
  })
}
