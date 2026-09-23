/**
 * Raw zip splicing for lazy media (Node only): entries are copied between
 * archives as stored bytes, never inflated or re-deflated, so a 1 GB document
 * costs the reads of its XML parts and nothing else.
 */
import { open, type FileHandle } from 'node:fs/promises'
import { crc32, inflateRawSync } from 'node:zlib'
import {
  LAZY_MEDIA_PLACEHOLDER_BYTES,
  isLazyMediaPart,
  lazyMediaHashOf,
  lazyMediaPlaceholder,
} from './lazy-media'

export interface ZipEntry {
  name: string
  nameBytes: Buffer
  flags: number
  method: number
  time: number
  date: number
  crc: number
  csize: number
  usize: number
  verMade: number
  verNeed: number
  intAttr: number
  extAttr: number
  /** first byte of the (compressed) data inside the archive */
  dataOffset: number
}

export interface ByteSource {
  size: number
  read(offset: number, length: number): Promise<Buffer>
}

export interface ZipFile extends ByteSource {
  entries: Map<string, ZipEntry>
  close(): Promise<void>
}

const SIG_LOCAL = 0x04034b50
const SIG_CENTRAL = 0x02014b50
const SIG_EOCD = 0x06054b50
const SIG_ZIP64_LOCATOR = 0x07064b50
const FLAG_ENCRYPTED = 0x1
const FLAG_DATA_DESCRIPTOR = 0x8
const FLAG_UTF8 = 0x800

export function bufferSource(buf: Buffer): ByteSource {
  return {
    size: buf.length,
    read: async (offset, length) => buf.subarray(offset, offset + length),
  }
}

function fileSource(fh: FileHandle, size: number): ByteSource {
  return {
    size,
    async read(offset, length) {
      const buf = Buffer.alloc(length)
      let done = 0
      while (done < length) {
        const { bytesRead } = await fh.read(buf, done, length - done, offset + done)
        if (bytesRead === 0) throw new Error('zip: unexpected end of file')
        done += bytesRead
      }
      return buf
    },
  }
}

export async function readZipEntries(src: ByteSource): Promise<ZipEntry[]> {
  const tailLen = Math.min(src.size, 22 + 0xffff)
  const tail = await src.read(src.size - tailLen, tailLen)
  let eocd = -1
  for (let i = tail.length - 22; i >= 0; i--) {
    if (tail.readUInt32LE(i) === SIG_EOCD) {
      eocd = i
      break
    }
  }
  if (eocd < 0) throw new Error('zip: end of central directory not found')
  if (eocd >= 20 && tail.readUInt32LE(eocd - 20) === SIG_ZIP64_LOCATOR) {
    throw new Error('zip: zip64 archives are not supported')
  }
  const count = tail.readUInt16LE(eocd + 10)
  const cdSize = tail.readUInt32LE(eocd + 12)
  const cdOffset = tail.readUInt32LE(eocd + 16)
  if (count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
    throw new Error('zip: zip64 archives are not supported')
  }
  if (cdOffset > src.size || cdSize > src.size || cdOffset + cdSize > src.size) {
    throw new Error('zip: corrupt central directory')
  }
  const cd = await src.read(cdOffset, cdSize)
  if (cd.length < cdSize) {
    throw new Error('zip: corrupt central directory')
  }
  const entries: ZipEntry[] = []
  let pos = 0
  for (let i = 0; i < count; i++) {
    if (pos + 46 > cd.length || cd.readUInt32LE(pos) !== SIG_CENTRAL) {
      throw new Error('zip: corrupt central directory')
    }
    const flags = cd.readUInt16LE(pos + 8)
    if (flags & FLAG_ENCRYPTED) throw new Error('zip: encrypted entries are not supported')
    const nameLen = cd.readUInt16LE(pos + 28)
    const extraLen = cd.readUInt16LE(pos + 30)
    const commentLen = cd.readUInt16LE(pos + 32)
    const nameBytes = Buffer.from(cd.subarray(pos + 46, pos + 46 + nameLen))
    const csize = cd.readUInt32LE(pos + 20)
    const usize = cd.readUInt32LE(pos + 24)
    const localOffset = cd.readUInt32LE(pos + 42)
    if (csize === 0xffffffff || usize === 0xffffffff || localOffset === 0xffffffff) {
      throw new Error('zip: zip64 archives are not supported')
    }
    entries.push({
      name: nameBytes.toString(flags & FLAG_UTF8 ? 'utf8' : 'latin1'),
      nameBytes,
      flags,
      method: cd.readUInt16LE(pos + 10),
      time: cd.readUInt16LE(pos + 12),
      date: cd.readUInt16LE(pos + 14),
      crc: cd.readUInt32LE(pos + 16),
      csize,
      usize,
      verMade: cd.readUInt16LE(pos + 4),
      verNeed: cd.readUInt16LE(pos + 6),
      intAttr: cd.readUInt16LE(pos + 36),
      extAttr: cd.readUInt32LE(pos + 38),
      dataOffset: localOffset,
    })
    pos += 46 + nameLen + extraLen + commentLen
  }
  // the local header's name/extra lengths may differ from the central copy
  for (const e of entries) {
    const local = await src.read(e.dataOffset, 30)
    if (local.length < 30 || local.readUInt32LE(0) !== SIG_LOCAL)
      throw new Error(`zip: corrupt local header for ${e.name}`)
    e.dataOffset += 30 + local.readUInt16LE(26) + local.readUInt16LE(28)
    // csize is the archive's own word about itself; readers below allocate
    // from it before touching the file, so a declared length past the end
    // (or past int32 — fs.read aborts the process on one) must never reach them.
    if (e.dataOffset > src.size || e.csize > src.size - e.dataOffset)
      throw new Error(
        `zip: entry ${e.name} declares ${e.csize} bytes at ${e.dataOffset}, ` +
          `outside the ${src.size}-byte archive`,
      )
  }
  return entries
}

/**
 * Opens a file-backed archive. The caller owns the handle and must call
 * close() in a finally block. The handle is closed automatically when
 * opening fails.
 */
export async function openZipFile(path: string): Promise<ZipFile> {
  const fh = await open(path, 'r')
  try {
    const { size } = await fh.stat()
    const src = fileSource(fh, size)
    const entries = new Map<string, ZipEntry>()
    for (const e of await readZipEntries(src)) entries.set(e.name, e)
    return { ...src, entries, close: () => fh.close() }
  } catch (err) {
    await fh.close()
    throw err
  }
}

interface OutEntry {
  meta: Omit<ZipEntry, 'dataOffset'>
  data: Buffer
}

/** Writes a plain (non-zip64) archive: data descriptors are folded into the
 *  headers and extra fields dropped, so every entry is self-describing. */
export function writeZip(entries: OutEntry[]): Buffer {
  if (entries.length >= 0xffff) throw new Error('zip: too many entries for a plain archive')
  const chunks: Buffer[] = []
  const centrals: Buffer[] = []
  let offset = 0
  for (const { meta, data } of entries) {
    const flags = meta.flags & ~FLAG_DATA_DESCRIPTOR
    const local = Buffer.alloc(30)
    local.writeUInt32LE(SIG_LOCAL, 0)
    local.writeUInt16LE(meta.verNeed, 4)
    local.writeUInt16LE(flags, 6)
    local.writeUInt16LE(meta.method, 8)
    local.writeUInt16LE(meta.time, 10)
    local.writeUInt16LE(meta.date, 12)
    local.writeUInt32LE(meta.crc, 14)
    local.writeUInt32LE(meta.csize, 16)
    local.writeUInt32LE(meta.usize, 20)
    local.writeUInt16LE(meta.nameBytes.length, 26)
    local.writeUInt16LE(0, 28)
    const central = Buffer.alloc(46)
    central.writeUInt32LE(SIG_CENTRAL, 0)
    central.writeUInt16LE(meta.verMade, 4)
    central.writeUInt16LE(meta.verNeed, 6)
    central.writeUInt16LE(flags, 8)
    central.writeUInt16LE(meta.method, 10)
    central.writeUInt16LE(meta.time, 12)
    central.writeUInt16LE(meta.date, 14)
    central.writeUInt32LE(meta.crc, 16)
    central.writeUInt32LE(meta.csize, 20)
    central.writeUInt32LE(meta.usize, 24)
    central.writeUInt16LE(meta.nameBytes.length, 28)
    central.writeUInt16LE(0, 30)
    central.writeUInt16LE(0, 32)
    central.writeUInt16LE(0, 34)
    central.writeUInt16LE(meta.intAttr, 36)
    central.writeUInt32LE(meta.extAttr, 38)
    central.writeUInt32LE(offset, 42)
    chunks.push(local, meta.nameBytes, data)
    centrals.push(central, meta.nameBytes)
    offset += 30 + meta.nameBytes.length + data.length
  }
  const cdSize = centrals.reduce((n, b) => n + b.length, 0)
  if (offset + cdSize > 0xffffffff) throw new Error('zip: archive too large for a plain archive')
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(SIG_EOCD, 0)
  eocd.writeUInt16LE(entries.length, 8)
  eocd.writeUInt16LE(entries.length, 10)
  eocd.writeUInt32LE(cdSize, 12)
  eocd.writeUInt32LE(offset, 16)
  return Buffer.concat([...chunks, ...centrals, eocd])
}

export interface SlimDocx {
  bytes: Buffer
  /** stripped part paths */
  lazyParts: string[]
}

/**
 * Copy of the archive with its lazy-media parts replaced by placeholders
 * naming `hash`; null when less than `minBytes` of media would be stripped.
 * Does not close zip; the caller retains ownership of the handle.
 */
export async function slimDocx(
  zip: ZipFile,
  hash: string,
  minBytes: number,
): Promise<SlimDocx | null> {
  const lazy = [...zip.entries.values()].filter((e) => isLazyMediaPart(e.name) && e.usize > 0)
  if (lazy.reduce((n, e) => n + e.usize, 0) < minBytes) return null
  const placeholder = Buffer.from(lazyMediaPlaceholder(hash))
  const placeholderCrc = crc32(placeholder)
  const out: OutEntry[] = []
  for (const e of zip.entries.values()) {
    if (isLazyMediaPart(e.name) && e.usize > 0) {
      out.push({
        meta: {
          ...e,
          method: 0,
          crc: placeholderCrc,
          csize: placeholder.length,
          usize: placeholder.length,
        },
        data: placeholder,
      })
    } else {
      out.push({ meta: e, data: await zip.read(e.dataOffset, e.csize) })
    }
  }
  return { bytes: writeZip(out), lazyParts: lazy.map((e) => e.name) }
}

function inflated(e: ZipEntry, raw: Buffer): Buffer {
  if (e.method === 0) return raw
  if (e.method === 8) return inflateRawSync(raw)
  throw new Error(`zip: unsupported compression method ${e.method} for ${e.name}`)
}

/** the placeholder hash when the entry is one, else null */
async function placeholderHash(src: ByteSource, e: ZipEntry, raw?: Buffer): Promise<string | null> {
  if (e.usize !== LAZY_MEDIA_PLACEHOLDER_BYTES) return null
  return lazyMediaHashOf(inflated(e, raw ?? (await src.read(e.dataOffset, e.csize))))
}

/** hashes named by the placeholders in a saved archive (empty when none) */
export async function lazyMediaHashesIn(bytes: Buffer): Promise<Set<string>> {
  const src = bufferSource(bytes)
  const hashes = new Set<string>()
  for (const e of await readZipEntries(src)) {
    const hash = await placeholderHash(src, e)
    if (hash) hashes.add(hash)
  }
  return hashes
}

/**
 * Saved archive with every placeholder swapped back for the original part
 * bytes; the input is returned as-is when it holds no placeholders.
 *
 * Takes ownership of every ZipFile returned by sourceFor and closes each
 * one before returning or throwing, including on error paths. Provide a
 * fresh handle per hash (for example `() => openZipFile(path)`) and do not
 * reuse a handle after the call.
 */
export async function materializeDocx(
  bytes: Buffer,
  sourceFor: (hash: string) => Promise<ZipFile | null>,
): Promise<Buffer> {
  const src = bufferSource(bytes)
  const entries = await readZipEntries(src)
  const out: OutEntry[] = []
  let swapped = false
  const cache = new Map<string, ZipFile | null>()
  const owned = new Set<ZipFile>()
  try {
    for (const e of entries) {
      const raw = await src.read(e.dataOffset, e.csize)
      const hash = await placeholderHash(src, e, raw)
      if (!hash) {
        out.push({ meta: e, data: raw })
        continue
      }
      let source = cache.get(hash)
      if (source === undefined) {
        source = await sourceFor(hash)
        cache.set(hash, source)
        if (source) owned.add(source)
      }
      const original = source?.entries.get(e.name)
      if (!source || !original) throw new Error(`lazy media source unavailable for ${e.name}`)
      out.push({
        meta: { ...original, nameBytes: e.nameBytes },
        data: await source.read(original.dataOffset, original.csize),
      })
      swapped = true
    }
    return swapped ? writeZip(out) : bytes
  } finally {
    for (const source of owned) {
      try {
        await source.close()
      } catch {
        // ignore close errors so the original result is preserved
      }
    }
  }
}
