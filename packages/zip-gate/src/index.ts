/**
 * Metered-inflation zip bomb gate, shared by the docx engine, the pptx engine
 * and the attachment parsers (file-parse).
 *
 * A part may declare any uncompressed size it likes, so a gate that reads the
 * central directory's declared sizes is advisory at best: 600 MB of payload
 * declaring 300 bytes passes it, and the first genuine inflate then pays for
 * the whole payload in one allocation (measured in #759: 794 MB and 3.12 GB of
 * RSS from sub-megabyte inputs; a later re-measurement agreed).
 *
 * `assertZipInflatesWithinLimits` inflates each part through a stream that is
 * cancelled one byte past what the part claims: an honest part completes at its
 * own size and a liar is caught at `declared + 1` bytes, so the cost of the
 * attack tracks the claim rather than the payload the claim hides. Stored parts
 * (every image in a real docx/pptx) are verified by length without inflating at
 * all, and summing verified lengths makes the total cap honest.
 *
 * Deliberately free of Node builtins and of jszip: every caller reaches either
 * a renderer bundle (where `node:*` imports fail Vite's externalized stubs) or
 * an engine that owns its own zip reader. Run `tsc --noEmit` in this package
 * with only the DOM lib to keep that honest.
 */

import { decompressionStream, streamBytes } from './byte-stream'

export { decompressionStream, streamBytes } from './byte-stream'

/** What the gates enforce; `DEFAULT_ZIP_LIMITS` is the default policy. */
export interface ZipLimits {
  maxParts: number
  maxPartBytes: number
  maxTotalBytes: number
}

/** Part-count / per-part / total caps shared by the docx and pptx engines. */
export const DEFAULT_ZIP_LIMITS = {
  maxParts: 10000,
  maxPartBytes: 512 * 1024 * 1024,
  maxTotalBytes: 1.5 * 1024 * 1024 * 1024,
} as const

/** What the declared-size pass needs to know about one part. */
export interface DeclaredPart {
  name: string
  usize: number
}

/**
 * The part-count / per-part / total declared-size gate.
 *
 * Split out of the JSZip-backed `assertZipWithinLimits` in the engines so a
 * raw entry reader can apply the same limits before it copies anything.
 * Advisory only — it costs nothing when the archive tells the truth and
 * nothing at all when it lies; `assertZipInflatesWithinLimits` is the gate
 * that holds against a forged declaration.
 */
export function assertDeclaredSizesWithinLimits(
  parts: readonly DeclaredPart[],
  limits: ZipLimits = DEFAULT_ZIP_LIMITS,
): void {
  if (parts.length > limits.maxParts) {
    throw new Error(`zip rejected: ${parts.length} parts exceeds the ${limits.maxParts} limit`)
  }
  let total = 0
  for (const part of parts) {
    if (part.usize > limits.maxPartBytes) {
      throw new Error(
        `zip rejected: part ${part.name} declares ${part.usize} uncompressed bytes ` +
          `(limit ${limits.maxPartBytes})`,
      )
    }
    if (part.usize > 0) total += part.usize
  }
  if (total > limits.maxTotalBytes) {
    throw new Error(
      `zip rejected: total uncompressed size ${total} exceeds the ` +
        `${limits.maxTotalBytes} limit`,
    )
  }
}

/** Everything the metered gate needs to know about one part. */
interface ScannedPart {
  name: string
  method: number
  csize: number
  usize: number
  /** first byte of the (compressed) data inside the archive */
  dataStart: number
}

const EOCD_SIG = 0x06054b50
const CENTRAL_SIG = 0x02014b50
const LOCAL_SIG = 0x04034b50
const ZIP64_LOCATOR_SIG = 0x07064b50
const FLAG_ENCRYPTED = 0x1

/**
 * The gate that actually holds against a forged central directory.
 *
 * Inflating one byte past what a part *claims* closes the advisory gap: an
 * honest part completes at its declared length and a liar is caught at
 * `declared + 1` bytes. An oversized *declaration* is still rejected before
 * anything is inflated, which is what keeps honest bombs at ~0 ms.
 */
export async function assertZipInflatesWithinLimits(
  bytes: Uint8Array,
  limits: ZipLimits = DEFAULT_ZIP_LIMITS,
): Promise<void> {
  const parts = scanParts(bytes)
  if (parts.length > limits.maxParts) {
    throw new Error(`zip rejected: ${parts.length} parts exceeds the ${limits.maxParts} limit`)
  }
  let total = 0
  for (const part of parts) {
    if (part.usize > limits.maxPartBytes) {
      throw new Error(
        `zip rejected: part ${part.name} declares ${part.usize} uncompressed bytes ` +
          `(limit ${limits.maxPartBytes})`,
      )
    }
    total += await verifiedSize(bytes, part)
    if (total > limits.maxTotalBytes) {
      throw new Error(
        `zip rejected: total uncompressed size ${total} exceeds the ` +
          `${limits.maxTotalBytes} limit`,
      )
    }
  }
}

/**
 * Reads the central directory of an archive already in memory.
 *
 * Deliberately local to this module and free of Node builtins: the engines'
 * renderer bundles reach this file, and richer readers — built on `node:fs`,
 * `node:zlib` and `Buffer` — cannot be tree-shaken out of those bundles once
 * something actually calls into them, which fails the browser build on Vite's
 * externalized stubs. Only the five fields the gate uses are read, and every
 * offset is bounds-checked against the file before it is touched.
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
 * payload nor a single large allocation is ever materialised. Same shape as
 * the gzip cap in the docx engine's metafile module, which is there for the
 * same renderer reason.
 */
async function inflateRawBounded(raw: Uint8Array, part: ScannedPart): Promise<number> {
  let seen = 0
  try {
    const reader = streamBytes(raw).pipeThrough(decompressionStream('deflate-raw')).getReader()
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
    if (err instanceof Error && err.message.startsWith('zip rejected:')) throw err
    // DecompressionStream rejects with an empty-message TypeError on bad data.
    const message = err instanceof Error ? err.message || err.name : String(err)
    throw new Error(`zip rejected: part ${part.name} cannot be inflated: ${message}`, {
      cause: err,
    })
  }
  return seen
}

/** Claims less than the bytes deliver: that gap is exactly where bombs live. */
function underDeclared(part: ScannedPart, cause?: unknown): Error {
  const error = new Error(
    `zip rejected: part ${part.name} declares ${part.usize} uncompressed bytes ` +
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
