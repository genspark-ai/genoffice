/**
 * Minimal dependency-free PNG encoder (RGBA8 → PNG bytes). Uses zlib "stored"
 * deflate blocks — no compression here; the docx container is a zip, so the
 * final file gets deflated anyway. Correctness over size for P1.
 */

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

function crc32(bytes: Uint8Array, start: number, end: number): number {
  let crc = 0xffffffff
  for (let i = start; i < end; i++) crc = CRC_TABLE[(crc ^ bytes[i]!) & 0xff]! ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

function adler32(bytes: Uint8Array): number {
  let a = 1
  let b = 0
  for (let i = 0; i < bytes.length; i++) {
    a = (a + bytes[i]!) % 65521
    b = (b + a) % 65521
  }
  return ((b << 16) | a) >>> 0
}

function writeU32(out: Uint8Array, pos: number, value: number): void {
  out[pos] = (value >>> 24) & 0xff
  out[pos + 1] = (value >>> 16) & 0xff
  out[pos + 2] = (value >>> 8) & 0xff
  out[pos + 3] = value & 0xff
}

/** length + type + data + crc(type+data) */
function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length)
  writeU32(out, 0, data.length)
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i)
  out.set(data, 8)
  writeU32(out, 8 + data.length, crc32(out, 4, 8 + data.length))
  return out
}

/** raw scanlines (filter byte 0 per row) wrapped in a stored-block zlib stream */
function zlibStored(raw: Uint8Array): Uint8Array {
  const MAX_BLOCK = 65535
  const blocks = Math.max(1, Math.ceil(raw.length / MAX_BLOCK))
  const out = new Uint8Array(2 + raw.length + blocks * 5 + 4)
  out[0] = 0x78 // CMF: deflate, 32k window
  out[1] = 0x01 // FLG: no dict, fastest
  let pos = 2
  for (let i = 0; i < blocks; i++) {
    const start = i * MAX_BLOCK
    const len = Math.min(MAX_BLOCK, raw.length - start)
    out[pos] = i === blocks - 1 ? 1 : 0 // BFINAL
    out[pos + 1] = len & 0xff
    out[pos + 2] = (len >>> 8) & 0xff
    out[pos + 3] = ~len & 0xff
    out[pos + 4] = (~len >>> 8) & 0xff
    out.set(raw.subarray(start, start + len), pos + 5)
    pos += 5 + len
  }
  writeU32(out, pos, adler32(raw))
  return out
}

/** Encode an RGBA8 buffer (4 bytes/pixel, tightly packed) as a PNG file. */
export function encodeRgbaPng(rgba: Uint8Array, width: number, height: number): Uint8Array {
  if (rgba.length !== width * height * 4) {
    throw new Error(`rgba buffer is ${rgba.length} bytes, expected ${width * height * 4}`)
  }
  const rowBytes = width * 4
  const raw = new Uint8Array(height * (1 + rowBytes))
  for (let y = 0; y < height; y++) {
    // filter type 0 (None) prepended to each scanline
    raw.set(rgba.subarray(y * rowBytes, (y + 1) * rowBytes), y * (1 + rowBytes) + 1)
  }

  const ihdr = new Uint8Array(13)
  writeU32(ihdr, 0, width)
  writeU32(ihdr, 4, height)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 6 // color type RGBA
  // compression/filter/interlace all 0

  const signature = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  const parts = [
    signature,
    chunk('IHDR', ihdr),
    chunk('IDAT', zlibStored(raw)),
    chunk('IEND', new Uint8Array(0)),
  ]
  const total = parts.reduce((n, p) => n + p.length, 0)
  const out = new Uint8Array(total)
  let pos = 0
  for (const p of parts) {
    out.set(p, pos)
    pos += p.length
  }
  return out
}
