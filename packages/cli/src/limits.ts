import { closeSync, fstatSync, openSync, readSync } from 'node:fs'
import { DOCX_ZIP_LIMITS } from '@genoffice/docx-engine'
import { PPTX_ZIP_LIMITS } from '@genoffice/pptx-engine'
import { XLSX_ZIP_LIMITS } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { CliError, EXIT } from './result'

export interface ZipDirectoryEntry {
  name: string
  compressed: number
  uncompressed: number
}

interface PackageLimits {
  maxParts: number
  maxPartBytes: number
  maxTotalBytes: number
}

/** Per-format ceilings are the engines' own, so the CLI never refuses what the engine would open. */
const LIMITS: Record<string, PackageLimits> = {
  docx: DOCX_ZIP_LIMITS,
  pptx: PPTX_ZIP_LIMITS,
  xlsx: { ...XLSX_ZIP_LIMITS, maxPartBytes: XLSX_ZIP_LIMITS.maxTotalBytes },
}

const FAMILY: Record<string, keyof typeof LIMITS> = {
  docx: 'docx',
  docm: 'docx',
  dotx: 'docx',
  dotm: 'docx',
  xlsx: 'xlsx',
  xlsm: 'xlsx',
  xltx: 'xlsx',
  xltm: 'xlsx',
  pptx: 'pptx',
  pptm: 'pptx',
  potx: 'pptx',
  potm: 'pptx',
  ppsx: 'pptx',
}

/** Real OOXML parts deflate well under 100:1; a bomb declares thousands to one. */
export const MAX_COMPRESSION_RATIO = 200
const RATIO_FLOOR_BYTES = 64 * 1024 * 1024

const EOCD_SIG = 0x06054b50
const EOCD64_LOCATOR_SIG = 0x07064b50
const EOCD64_SIG = 0x06064b50
const CENTRAL_SIG = 0x02014b50
const ZIP64_EXTRA_ID = 0x0001
const MAX_TAIL = 22 + 0xffff + 20

function readAt(fd: number, offset: number, length: number): Buffer {
  const buf = Buffer.alloc(length)
  let done = 0
  while (done < length) {
    const n = readSync(fd, buf, done, length - done, offset + done)
    if (n === 0) break
    done += n
  }
  return done === length ? buf : buf.subarray(0, done)
}

function u64(buf: Buffer, at: number): number {
  return Number(buf.readBigUInt64LE(at))
}

export interface ZipDirectory {
  /** entries the end record declares; `entries` is left empty when it is over `maxEntries` */
  count: number
  entries: ZipDirectoryEntry[]
}

/** A legitimate 10 000-part directory is a couple of MB; anything near this is hostile. */
const MAX_DIRECTORY_BYTES = 64 * 1024 * 1024

/**
 * Entries declared in the central directory, read straight from the file
 * without inflating anything. Null when the file does not end in a zip
 * end-of-central-directory record (the engine reports that in its own words).
 * The directory itself is not read when it declares more than `maxEntries`
 * parts or is larger than a sane directory can be, so a hostile count cannot
 * make the guard allocate.
 */
export function readZipDirectory(path: string, maxEntries = Infinity): ZipDirectory | null {
  const fd = openSync(path, 'r')
  try {
    const size = fstatSync(fd).size
    const tailStart = Math.max(0, size - MAX_TAIL)
    const tail = readAt(fd, tailStart, size - tailStart)
    let eocd = -1
    for (let i = tail.length - 22; i >= 0; i--) {
      if (tail.readUInt32LE(i) === EOCD_SIG) {
        eocd = i
        break
      }
    }
    if (eocd < 0) return null
    let count = tail.readUInt16LE(eocd + 10)
    let cdSize = tail.readUInt32LE(eocd + 12)
    let cdOffset = tail.readUInt32LE(eocd + 16)
    // JSZip's rule: the zip64 record is consulted only when the end record overflowed
    // (an 0xFFFF / 0xFFFFFFFF sentinel); a locator in front of a plain end record is
    // ignored, so the guard reads the same directory the engine will.
    const overflowed =
      tail.readUInt16LE(eocd + 4) === 0xffff ||
      tail.readUInt16LE(eocd + 6) === 0xffff ||
      tail.readUInt16LE(eocd + 8) === 0xffff ||
      count === 0xffff ||
      cdSize === 0xffffffff ||
      cdOffset === 0xffffffff
    const loc = eocd - 20
    if (overflowed && loc >= 0 && tail.readUInt32LE(loc) === EOCD64_LOCATOR_SIG) {
      const rec = readAt(fd, u64(tail, loc + 8), 56)
      if (rec.length === 56 && rec.readUInt32LE(0) === EOCD64_SIG) {
        count = u64(rec, 32)
        cdSize = u64(rec, 40)
        cdOffset = u64(rec, 48)
      }
    }
    if (count > maxEntries || cdSize > MAX_DIRECTORY_BYTES) return { count, entries: [] }
    if (cdOffset + cdSize > size) return null
    const cd = readAt(fd, cdOffset, cdSize)
    const entries: ZipDirectoryEntry[] = []
    let p = 0
    for (let i = 0; i < count && p + 46 <= cd.length; i++) {
      if (cd.readUInt32LE(p) !== CENTRAL_SIG) break
      let compressed = cd.readUInt32LE(p + 20)
      let uncompressed = cd.readUInt32LE(p + 24)
      const nameLen = cd.readUInt16LE(p + 28)
      const extraLen = cd.readUInt16LE(p + 30)
      const commentLen = cd.readUInt16LE(p + 32)
      const name = cd.subarray(p + 46, p + 46 + nameLen).toString('utf8')
      let q = p + 46 + nameLen
      const extraEnd = Math.min(q + extraLen, cd.length)
      while (q + 4 <= extraEnd) {
        const id = cd.readUInt16LE(q)
        const len = cd.readUInt16LE(q + 2)
        if (id === ZIP64_EXTRA_ID) {
          let r = q + 4
          if (uncompressed === 0xffffffff && r + 8 <= extraEnd) {
            uncompressed = u64(cd, r)
            r += 8
          }
          if (compressed === 0xffffffff && r + 8 <= extraEnd) compressed = u64(cd, r)
        }
        q += 4 + len
      }
      if (!name.endsWith('/')) entries.push({ name, compressed, uncompressed })
      p += 46 + nameLen + extraLen + commentLen
    }
    return { count, entries }
  } finally {
    closeSync(fd)
  }
}

/** Same rule as the xlsx gateway: only a `..` that walks out of the root (or a NUL) is an escape; leading slashes and backslashes are producer quirks it folds away. */
function unsafeName(name: string): boolean {
  if (name.includes('\0')) return true
  let depth = 0
  for (const seg of name.split(/[/\\]/)) {
    if (seg === '' || seg === '.') continue
    if (seg === '..') {
      if (depth === 0) return true
      depth--
    } else depth++
  }
  return false
}

function refuse(message: string, detail: Record<string, unknown>, suggestion: string): CliError {
  return new CliError(EXIT.file, message, detail, { reason: 'resource_limit', suggestion })
}

/**
 * Refuse a package the engine would choke on before any part is inflated:
 * too many parts, oversized parts, absurd compression ratios or entry names
 * that escape the archive. Non-OOXML extensions and non-zip files pass.
 */
export function assertPackageWithinLimits(path: string, ext: string): void {
  const family = FAMILY[ext]
  if (!family) return
  const limits = LIMITS[family]
  const dir = readZipDirectory(path, limits.maxParts)
  if (!dir) return
  if (dir.count > limits.maxParts || !dir.entries.length) {
    throw refuse(
      `${path}: ${dir.count} zip entries exceeds the ${limits.maxParts} limit`,
      { entries: dir.count, limit: limits.maxParts },
      'reduce the number of parts (embedded media, sheets, slides) in the package',
    )
  }
  const { entries } = dir
  const escaping = entries.find((e) => unsafeName(e.name))
  if (escaping) {
    throw refuse(
      `${path}: zip entry "${escaping.name}" escapes the package`,
      { entry: escaping.name },
      'repair the file in Office or re-save it; genoffice does not open packages with unsafe entry paths',
    )
  }
  let total = 0
  let totalCompressed = 0
  for (const e of entries) {
    if (e.uncompressed > limits.maxPartBytes) {
      throw refuse(
        `${path}: part ${e.name} declares ${e.uncompressed} uncompressed bytes (limit ${limits.maxPartBytes})`,
        { entry: e.name, bytes: e.uncompressed, limit: limits.maxPartBytes },
        'remove or shrink the oversized part; genoffice refuses packages over detail.limit bytes',
      )
    }
    if (
      e.uncompressed > RATIO_FLOOR_BYTES &&
      e.uncompressed / Math.max(1, e.compressed) > MAX_COMPRESSION_RATIO
    ) {
      throw refuse(
        `${path}: part ${e.name} inflates ${e.compressed} → ${e.uncompressed} bytes, beyond the ${MAX_COMPRESSION_RATIO}:1 limit`,
        {
          entry: e.name,
          compressed: e.compressed,
          bytes: e.uncompressed,
          limit: MAX_COMPRESSION_RATIO,
        },
        'this looks like a decompression bomb; open the file in Office to verify it before retrying',
      )
    }
    total += e.uncompressed
    totalCompressed += e.compressed
  }
  if (total > limits.maxTotalBytes) {
    throw refuse(
      `${path}: total uncompressed size ${total} exceeds the ${limits.maxTotalBytes} limit`,
      { bytes: total, limit: limits.maxTotalBytes },
      'split the document or remove embedded media; genoffice refuses packages over detail.limit bytes',
    )
  }
  if (total > RATIO_FLOOR_BYTES && total / Math.max(1, totalCompressed) > MAX_COMPRESSION_RATIO) {
    throw refuse(
      `${path}: package inflates ${totalCompressed} → ${total} bytes, beyond the ${MAX_COMPRESSION_RATIO}:1 limit`,
      { compressed: totalCompressed, bytes: total, limit: MAX_COMPRESSION_RATIO },
      'this looks like a decompression bomb; open the file in Office to verify it before retrying',
    )
  }
}
