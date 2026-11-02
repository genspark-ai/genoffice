/**
 * BMP → PNG. Hangul documents embed scanned seals, logos and form graphics as
 * Windows bitmaps (BinData/imageN.bmp), which OOXML consumers do not all
 * render and docx-engine does not accept. Uncompressed 1/4/8/24/32-bit DIBs
 * (BI_RGB, and BI_BITFIELDS with the standard masks) are re-encoded as 8-bit
 * RGB(A) PNGs; RLE-compressed or exotic headers return null so the caller can
 * count the picture as dropped instead of emitting a broken one.
 */

import { deflateSync } from 'node:zlib'

const MAX_PIXELS = 40_000_000

export function bmpToPng(bytes: Uint8Array): Uint8Array | null {
  if (bytes.length < 54 || bytes[0] !== 0x42 || bytes[1] !== 0x4d) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const pixelOffset = view.getUint32(10, true)
  const headerSize = view.getUint32(14, true)
  if (headerSize < 40) return null // OS/2 BITMAPCOREHEADER: not worth carrying
  const width = view.getInt32(18, true)
  const rawHeight = view.getInt32(22, true)
  const bitCount = view.getUint16(28, true)
  const compression = view.getUint32(30, true)
  const height = Math.abs(rawHeight)
  const topDown = rawHeight < 0
  if (width <= 0 || height <= 0 || width * height > MAX_PIXELS) return null
  if (compression !== 0 && !(compression === 3 && (bitCount === 32 || bitCount === 16))) return null
  if (![1, 4, 8, 16, 24, 32].includes(bitCount)) return null

  const stride = Math.floor((bitCount * width + 31) / 32) * 4
  if (pixelOffset + stride * height > bytes.length) return null

  let palette: Uint8Array | null = null
  if (bitCount <= 8) {
    // biClrUsed is the file's word: never more entries than the bit count indexes
    const colors = Math.min(view.getUint32(46, true) || 1 << bitCount, 1 << bitCount)
    const start = 14 + headerSize
    palette = new Uint8Array(colors * 3)
    for (let i = 0; i < colors && start + i * 4 + 2 < bytes.length; i++) {
      palette[i * 3] = bytes[start + i * 4 + 2]
      palette[i * 3 + 1] = bytes[start + i * 4 + 1]
      palette[i * 3 + 2] = bytes[start + i * 4]
    }
  }

  // 32-bit DIBs only carry alpha when some pixel actually uses it
  let hasAlpha = false
  if (bitCount === 32) {
    for (let y = 0; y < height && !hasAlpha; y++) {
      const row = pixelOffset + y * stride
      for (let x = 0; x < width; x++) {
        if (bytes[row + x * 4 + 3] !== 0) {
          hasAlpha = true
          break
        }
      }
    }
  }
  const channels = hasAlpha ? 4 : 3
  const raw = new Uint8Array((width * channels + 1) * height)
  for (let y = 0; y < height; y++) {
    const srcRow = pixelOffset + (topDown ? y : height - 1 - y) * stride
    let dst = y * (width * channels + 1)
    raw[dst++] = 0 // filter: none
    for (let x = 0; x < width; x++) {
      let r: number
      let g: number
      let b: number
      let a = 255
      if (bitCount === 24 || bitCount === 32) {
        const p = srcRow + x * (bitCount / 8)
        b = bytes[p]
        g = bytes[p + 1]
        r = bytes[p + 2]
        if (bitCount === 32) a = bytes[p + 3]
      } else if (bitCount === 16) {
        // 5-5-5 (BI_RGB) — BI_BITFIELDS 5-6-5 is rare enough to share the path
        const v = view.getUint16(srcRow + x * 2, true)
        r = ((v >> 10) & 0x1f) << 3
        g = ((v >> 5) & 0x1f) << 3
        b = (v & 0x1f) << 3
      } else {
        const bitIndex = x * bitCount
        const byte = bytes[srcRow + (bitIndex >> 3)]
        const shift = 8 - bitCount - (bitIndex & 7)
        const index = (byte >> shift) & ((1 << bitCount) - 1)
        r = palette![index * 3] ?? 0
        g = palette![index * 3 + 1] ?? 0
        b = palette![index * 3 + 2] ?? 0
      }
      raw[dst++] = r
      raw[dst++] = g
      raw[dst++] = b
      if (hasAlpha) raw[dst++] = a
    }
  }

  const ihdr = new Uint8Array(13)
  const ihdrView = new DataView(ihdr.buffer)
  ihdrView.setUint32(0, width)
  ihdrView.setUint32(4, height)
  ihdr[8] = 8
  ihdr[9] = hasAlpha ? 6 : 2
  return concat([
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', new Uint8Array(deflateSync(raw, { level: 6 }))),
    chunk('IEND', new Uint8Array(0)),
  ])
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

function crc32(parts: Uint8Array[]): number {
  let c = 0xffffffff
  for (const part of parts) {
    for (let i = 0; i < part.length; i++) c = CRC_TABLE[(c ^ part[i]) & 0xff] ^ (c >>> 8)
  }
  return (c ^ 0xffffffff) >>> 0
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = new TextEncoder().encode(type)
  const out = new Uint8Array(12 + data.length)
  const view = new DataView(out.buffer)
  view.setUint32(0, data.length)
  out.set(typeBytes, 4)
  out.set(data, 8)
  view.setUint32(8 + data.length, crc32([typeBytes, data]))
  return out
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}
