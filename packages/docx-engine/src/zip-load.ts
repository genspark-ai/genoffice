import JSZip from 'jszip'
import { needsOoxmlNormalization, normalizeOoxmlXml } from './ooxml-normalize'

const EOCD_SIG = 0x06054b50
const CENTRAL_SIG = 0x02014b50
const LOCAL_SIG = 0x04034b50
const ZIP64_LOCATOR_SIG = 0x07064b50
const UNICODE_PATH_ID = 0x7075
const FLAG_ENCRYPTED = 0x1

/**
 * Word resolves docx parts by the zip header file names and ignores Info-ZIP
 * Unicode Path (0x7075) extra fields; JSZip honors them, so a crc-valid but
 * conflicting field can shadow word/document.xml with another entry's bytes
 * (POI's unicode-path corpus). Blank the field id in the central directory so
 * JSZip falls back to the header names. Returns the input when nothing to do.
 */
function neutralizeUnicodePathFields(bytes: Uint8Array): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  const stop = Math.max(0, bytes.length - 22 - 0xffff)
  for (let i = bytes.length - 22; i >= stop; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocd = i
      break
    }
  }
  if (eocd < 0) return bytes
  const count = view.getUint16(eocd + 10, true)
  const cdOffset = view.getUint32(eocd + 16, true)
  if (count === 0xffff || cdOffset === 0xffffffff) return bytes // zip64: leave as-is
  let out: Uint8Array | null = null
  let p = cdOffset
  for (let i = 0; i < count; i++) {
    if (p + 46 > bytes.length || view.getUint32(p, true) !== CENTRAL_SIG) return out ?? bytes
    const nameLen = view.getUint16(p + 28, true)
    const extraLen = view.getUint16(p + 30, true)
    const commentLen = view.getUint16(p + 32, true)
    let q = p + 46 + nameLen
    const extraEnd = Math.min(q + extraLen, bytes.length)
    while (q + 4 <= extraEnd) {
      const fieldId = view.getUint16(q, true)
      const fieldLen = view.getUint16(q + 2, true)
      if (fieldId === UNICODE_PATH_ID) {
        out ??= new Uint8Array(bytes)
        out[q] = 0xff
        out[q + 1] = 0xff
      }
      q += 4 + fieldLen
    }
    p += 46 + nameLen + extraLen + commentLen
  }
  return out ?? bytes
}

/** Shared with the CLI's pre-open check so both layers accept the same files. */
export const DOCX_ZIP_LIMITS = {
  maxParts: 10000,
  maxPartBytes: 512 * 1024 * 1024,
  maxTotalBytes: 1.5 * 1024 * 1024 * 1024,
} as const
const MAX_ZIP_PARTS = DOCX_ZIP_LIMITS.maxParts
const MAX_PART_UNCOMPRESSED_BYTES = DOCX_ZIP_LIMITS.maxPartBytes
const MAX_TOTAL_UNCOMPRESSED_BYTES = DOCX_ZIP_LIMITS.maxTotalBytes

/**
 * Cheap fast path: reject zip bombs using the uncompressed sizes the central
 * directory declares (JSZip keeps them in the lazy `_data` compressed object).
 *
 * Advisory only — it costs nothing when the archive tells the truth and nothing
 * at all when it lies. `assertZipInflatesWithinLimits` below is the gate that
 * holds against a forged declaration.
 */
export function assertZipWithinLimits(zip: JSZip): void {
  const files = Object.values(zip.files).filter((f) => !f.dir)
  if (files.length > MAX_ZIP_PARTS) {
    throw new Error(`docx rejected: ${files.length} parts exceeds the ${MAX_ZIP_PARTS} limit`)
  }
  let total = 0
  for (const file of files) {
    const size =
      (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0
    if (size > MAX_PART_UNCOMPRESSED_BYTES) {
      throw new Error(
        `docx rejected: part ${file.name} declares ${size} uncompressed bytes ` +
          `(limit ${MAX_PART_UNCOMPRESSED_BYTES})`,
      )
    }
    if (size > 0) total += size
  }
  if (total > MAX_TOTAL_UNCOMPRESSED_BYTES) {
    throw new Error(
      `docx rejected: total uncompressed size ${total} exceeds the ` +
        `${MAX_TOTAL_UNCOMPRESSED_BYTES} limit`,
    )
  }
}

/** What `assertZipInflatesWithinLimits` enforces; `DOCX_ZIP_LIMITS` is the default. */
export interface ZipLimits {
  maxParts: number
  maxPartBytes: number
  maxTotalBytes: number
}

/** Everything the gate needs to know about one part. */
interface ScannedPart {
  name: string
  method: number
  csize: number
  usize: number
  /** first byte of the (compressed) data inside the archive */
  dataStart: number
}

/**
 * The gate that actually holds against a forged central directory.
 *
 * A part may declare any uncompressed size it likes, so the declared-size pass
 * above is walked straight past: 600 MB of payload declaring 300 bytes passes
 * it, and the first genuine inflate — `normalizeOoxmlParts` below, or any
 * caller's `file.async()` — then pays for the whole payload in one allocation.
 * Measured in #759: 794 MB and 3.12 GB of RSS from a 100 KB and a 272-byte input.
 *
 * Inflating one byte past what a part *claims* closes that: an honest part
 * completes at its declared length and a liar is caught at `declared + 1` bytes,
 * so the cost of the attack tracks the claim rather than the payload the claim
 * hides. The budget is therefore never the limit itself. An oversized
 * *declaration* is still rejected before anything is inflated, which is what
 * keeps honest bombs at ~0 ms, and stored parts (every image in a real docx)
 * are verified by length without inflating at all.
 *
 * Summing verified lengths also makes the total cap honest, where the declared
 * pass above can only add up claims.
 */
export async function assertZipInflatesWithinLimits(
  bytes: Uint8Array,
  limits: ZipLimits = DOCX_ZIP_LIMITS,
): Promise<void> {
  const parts = scanParts(bytes)
  if (parts.length > limits.maxParts) {
    throw new Error(`docx rejected: ${parts.length} parts exceeds the ${limits.maxParts} limit`)
  }
  let total = 0
  for (const part of parts) {
    if (part.usize > limits.maxPartBytes) {
      throw new Error(
        `docx rejected: part ${part.name} declares ${part.usize} uncompressed bytes ` +
          `(limit ${limits.maxPartBytes})`,
      )
    }
    total += await verifiedSize(bytes, part)
    if (total > limits.maxTotalBytes) {
      throw new Error(
        `docx rejected: total uncompressed size ${total} exceeds the ` +
          `${limits.maxTotalBytes} limit`,
      )
    }
  }
}

/**
 * Reads the central directory of an archive already in memory.
 *
 * Deliberately local to this module and free of Node builtins: the Docs renderer
 * bundles this file through `parseDocx`, and `zip-splice`'s richer reader — built
 * on `node:fs`, `node:zlib` and `Buffer` — cannot be tree-shaken out of that
 * bundle once something actually calls into it, which fails the browser build on
 * Vite's externalized stubs. Only the five fields the gate uses are read, and
 * every offset is bounds-checked against the file before it is touched.
 */
function scanParts(bytes: Uint8Array): ScannedPart[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const at = (i: number, width: number): number =>
    width === 4 ? view.getUint32(i, true) : view.getUint16(i, true)
  const end = bytes.byteLength - 22
  if (end < 0) throw new Error('zip: end of central directory not found')
  let eocd = -1
  for (let i = end; i >= Math.max(0, end - 0xffff); i--) {
    if (at(i, 4) === EOCD_SIG) {
      eocd = i
      break
    }
  }
  if (eocd < 0) throw new Error('zip: end of central directory not found')
  if (eocd >= 20 && at(eocd - 20, 4) === ZIP64_LOCATOR_SIG) {
    throw new Error('zip: zip64 archives are not supported')
  }
  const count = at(eocd + 10, 2)
  const cdSize = at(eocd + 12, 4)
  const cdOffset = at(eocd + 16, 4)
  if (count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
    throw new Error('zip: zip64 archives are not supported')
  }
  if (cdOffset > bytes.byteLength || cdSize > bytes.byteLength - cdOffset) {
    throw new Error('zip: corrupt central directory')
  }
  const parts: ScannedPart[] = []
  let pos = cdOffset
  for (let i = 0; i < count; i++) {
    if (pos + 46 > cdOffset + cdSize || at(pos, 4) !== CENTRAL_SIG) {
      throw new Error('zip: corrupt central directory')
    }
    const flags = at(pos + 8, 2)
    if (flags & FLAG_ENCRYPTED) throw new Error('zip: encrypted entries are not supported')
    const nameLen = at(pos + 28, 2)
    const extraLen = at(pos + 30, 2)
    const commentLen = at(pos + 32, 2)
    const nameStart = pos + 46
    if (nameStart + nameLen > cdOffset + cdSize) {
      throw new Error('zip: corrupt central directory')
    }
    const localOffset = at(pos + 42, 4)
    const rawName = bytes.subarray(nameStart, nameStart + nameLen)
    parts.push({
      // latin1 keeps byte-for-byte round-tripping of the non-UTF8 names Word and
      // JSZip accept; the string is only ever used to name a rejection.
      name: decodeLatin1(rawName),
      method: at(pos + 10, 2),
      csize: at(pos + 20, 4),
      usize: at(pos + 24, 4),
      dataStart: localOffset,
    })
    pos = nameStart + nameLen + extraLen + commentLen
  }
  // The local header's own lengths win: they may differ from the central copy.
  for (const part of parts) {
    if (part.dataStart + 30 > bytes.byteLength || at(part.dataStart, 4) !== LOCAL_SIG) {
      throw new Error(`zip: corrupt local header for ${part.name}`)
    }
    const start = part.dataStart + 30 + at(part.dataStart + 26, 2) + at(part.dataStart + 28, 2)
    if (start > bytes.byteLength || part.csize > bytes.byteLength - start) {
      throw new Error(
        `zip: entry ${part.name} declares ${part.csize} bytes at ${start}, ` +
          `outside the ${bytes.byteLength}-byte archive`,
      )
    }
    part.dataStart = start
  }
  // Directories carry no payload and JSZip's own gate does not count them.
  return parts.filter((part) => !part.name.endsWith('/'))
}

/** The part's real length, refusing to inflate past what it declared. */
async function verifiedSize(bytes: Uint8Array, part: ScannedPart): Promise<number> {
  const raw = bytes.subarray(part.dataStart, part.dataStart + part.csize)
  if (part.method === 0) {
    if (raw.length > part.usize) throw underDeclared(part)
    return raw.length
  }
  if (part.method !== 8) {
    throw new Error(`zip: unsupported compression method ${part.method} for ${part.name}`)
  }
  return inflateRawBounded(raw, part)
}

/**
 * Inflates through a stream and stops reading at the claim, so neither the
 * payload nor a single large allocation is ever materialised. Same shape as the
 * gzip cap in `metafile.ts`, which is here for the same renderer reason.
 */
async function inflateRawBounded(raw: Uint8Array, part: ScannedPart): Promise<number> {
  let seen = 0
  try {
    // Copy to a fresh ArrayBuffer-backed view: BlobPart rejects ArrayBufferLike
    // (metafile.ts carries the same note), and the cli tsconfig's DOM lib is
    // stricter about it than this package's own.
    const reader = new Blob([new Uint8Array(raw)])
      .stream()
      .pipeThrough(new DecompressionStream('deflate-raw'))
      .getReader()
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      seen += value.byteLength
      if (seen > part.usize) {
        await reader.cancel()
        throw underDeclared(part)
      }
    }
  } catch (err) {
    if (err instanceof Error && err.message.startsWith('docx rejected:')) throw err
    const message = (err as Error).message ?? String(err)
    throw new Error(`docx rejected: part ${part.name} cannot be inflated: ${message}`, {
      cause: err,
    })
  }
  return seen
}

/** Claims less than the bytes deliver: that gap is exactly where bombs live. */
function underDeclared(part: ScannedPart, cause?: unknown): Error {
  const error = new Error(
    `docx rejected: part ${part.name} declares ${part.usize} uncompressed bytes ` +
      'but inflates past that',
  )
  return cause === undefined ? error : Object.assign(error, { cause })
}

/** UTF-8 with the high bytes preserved, matching what JSZip does with names. */
function decodeLatin1(bytes: Uint8Array): string {
  let out = ''
  for (const byte of bytes) out += String.fromCharCode(byte)
  return out
}

// The gate parts carry the strict / non-canonical-prefix markers whenever the
// package needs normalizing (strict packages are strict in document.xml and
// their rels; prefix oddities live in document.xml).
const NORMALIZE_GATE_PARTS = ['word/document.xml', 'word/_rels/document.xml.rels', '_rels/.rels']

/**
 * ISO Strict OOXML and non-canonical namespace prefixes are normalized here,
 * at the single zip entry point, so parseDocx offsets and saveDocx patches
 * operate on identical part text (and saved packages come out transitional).
 */
async function normalizeOoxmlParts(zip: JSZip): Promise<void> {
  let needed = false
  for (const path of NORMALIZE_GATE_PARTS) {
    const file = zip.file(path)
    if (file && needsOoxmlNormalization(await file.async('string'))) {
      needed = true
      break
    }
  }
  if (!needed) return
  for (const file of Object.values(zip.files)) {
    if (file.dir || !/\.(xml|rels)$/i.test(file.name)) continue
    const xml = await file.async('string')
    if (needsOoxmlNormalization(xml)) zip.file(file.name, normalizeOoxmlXml(xml))
  }
}

/** Load a docx/zip resolving part names the way Word does. */
export async function loadDocxZip(bytes: Uint8Array): Promise<JSZip> {
  const prepared = neutralizeUnicodePathFields(bytes)
  // Metered inflation first: past this point no part is larger than a limit,
  // so the inflates normalizeOoxmlParts and every caller perform are bounded.
  await assertZipInflatesWithinLimits(prepared)
  const zip = await JSZip.loadAsync(prepared)
  assertZipWithinLimits(zip)
  await normalizeOoxmlParts(zip)
  return zip
}
