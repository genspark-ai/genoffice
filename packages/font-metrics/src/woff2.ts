/**
 * WOFF2 -> sfnt for metrics reads only. The table stream is one brotli block; the
 * untransformed tables (cmap, head, hhea, hmtx without the lsb transform, ...) come
 * out byte-identical, so a table directory over them serves the advance reader.
 * Transformed glyf / loca / hmtx are left out: outlines are not needed here, and a
 * transformed hmtx means the face has no plain advances to read.
 */
import { brotliDecompressSync } from 'node:zlib'

const KNOWN_TAGS = (
  'cmap head hhea hmtx maxp name OS/2 post cvt  fpgm glyf loca prep CFF  VORG EBDT EBLC gasp ' +
  'hdmx kern LTSH PCLT VDMX vhea vmtx BASE GDEF GPOS GSUB EBSC JSTF MATH CBDT CBLC COLR CPAL ' +
  'SVG  sbix acnt avar bdat bloc bsln cvar fdsc feat fmtx fvar gvar hsty just lcar mort morx ' +
  'opbd prop trak Zapf Silf Glat Gloc Feat Sill'
)
  .match(/.{4}(?: |$)/g)!
  .map((t) => t.slice(0, 4))

const WOFF2_SIGNATURE = 0x774f4632 // 'wOF2'
const TTC_FLAVOR = 0x74746366 // 'ttcf'

function readBase128(buf: Buffer, pos: number): [number, number] {
  let value = 0
  for (let i = 0; i < 5; i++) {
    const byte = buf[pos + i]
    if (byte === undefined) throw new Error('woff2: truncated base-128 integer')
    value = value * 128 + (byte & 0x7f)
    if ((byte & 0x80) === 0) return [value, pos + i + 1]
  }
  throw new Error('woff2: base-128 integer over 5 bytes')
}

export function isWoff2(buf: Buffer): boolean {
  return buf.length >= 48 && buf.readUInt32BE(0) === WOFF2_SIGNATURE
}

/** sfnt bytes holding the untransformed tables of a single-face WOFF2; null when not one */
export function woff2ToSfnt(data: Buffer): Buffer | null {
  if (!isWoff2(data)) return null
  const flavor = data.readUInt32BE(4)
  if (flavor === TTC_FLAVOR) return null
  const numTables = data.readUInt16BE(12)
  const compressedSize = data.readUInt32BE(20)
  let pos = 48
  const entries: Array<{ tag: string; len: number; transformed: boolean }> = []
  for (let t = 0; t < numTables; t++) {
    const flags = data[pos++]!
    let tag: string
    if ((flags & 0x3f) === 63) {
      tag = data.toString('latin1', pos, pos + 4)
      pos += 4
    } else tag = KNOWN_TAGS[flags & 0x3f]!
    const version = flags >> 6
    const transformed = tag === 'glyf' || tag === 'loca' ? version === 0 : version !== 0
    let len: number
    ;[len, pos] = readBase128(data, pos)
    if (transformed) [len, pos] = readBase128(data, pos)
    entries.push({ tag, len, transformed })
  }
  let stream: Buffer
  try {
    stream = brotliDecompressSync(data.subarray(pos, pos + compressedSize))
  } catch {
    return null
  }
  const kept: Array<{ tag: string; bytes: Buffer }> = []
  let off = 0
  for (const e of entries) {
    if (off + e.len > stream.length) return null
    if (!e.transformed) kept.push({ tag: e.tag, bytes: stream.subarray(off, off + e.len) })
    off += e.len
  }
  if (!kept.length) return null
  kept.sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0))
  const dirSize = 12 + 16 * kept.length
  let total = dirSize
  for (const k of kept) total += (k.bytes.length + 3) & ~3
  const out = Buffer.alloc(total)
  out.writeUInt32BE(flavor, 0)
  out.writeUInt16BE(kept.length, 4)
  let at = dirSize
  kept.forEach((k, i) => {
    out.write(k.tag, 12 + 16 * i, 4, 'latin1')
    out.writeUInt32BE(at, 12 + 16 * i + 8)
    out.writeUInt32BE(k.bytes.length, 12 + 16 * i + 12)
    k.bytes.copy(out, at)
    at += (k.bytes.length + 3) & ~3
  })
  return out
}
