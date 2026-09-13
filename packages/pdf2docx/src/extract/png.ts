/**
 * Minimal dependency-free PNG encoder (RGBA8 in → PNG bytes out). The pixel
 * buffer is analysed once and written in the cheapest colour type that
 * reproduces it exactly — indexed (≤256 distinct RGBA values, 1/2/4/8 bpp),
 * greyscale, greyscale+alpha, RGB, or RGBA — and the scanlines go through a
 * real deflate. A stencil-mask scan that used to leave as a 17 MB stored
 * RGBA stream is a few KB; office containers store media parts as-is, so
 * the PNG has to be small on its own.
 */
import { deflateSync } from 'node:zlib'

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

/** deflate level: 6 is zlib's default trade-off; scans and flat art compress
 * 1000:1 at any level, photos gain little above it */
const DEFLATE_LEVEL = 6
/** palettes beyond this many entries fall back to a truecolour/grey type */
const PALETTE_MAX = 256

const COLOR_GREY = 0
const COLOR_RGB = 2
const COLOR_INDEXED = 3
const COLOR_GREY_ALPHA = 4
const COLOR_RGBA = 6

interface Analysis {
  opaque: boolean
  grey: boolean
  /** distinct RGBA values as packed uint32 keys, insertion order; undefined
   * once the count passes PALETTE_MAX */
  palette: number[] | undefined
}

function analyse(rgba: Uint8Array): Analysis {
  let opaque = true
  let grey = true
  const seen = new Map<number, number>()
  let palette: number[] | undefined = []
  for (let i = 0; i < rgba.length; i += 4) {
    const r = rgba[i]!
    const g = rgba[i + 1]!
    const b = rgba[i + 2]!
    const a = rgba[i + 3]!
    if (a !== 255) opaque = false
    if (r !== g || g !== b) grey = false
    if (palette !== undefined) {
      const key = ((r << 24) | (g << 16) | (b << 8) | a) >>> 0
      if (!seen.has(key)) {
        if (seen.size >= PALETTE_MAX) palette = undefined
        else {
          seen.set(key, seen.size)
          palette.push(key)
        }
      }
    }
    if (!opaque && !grey && palette === undefined) break
  }
  return { opaque, grey, palette }
}

/** filter byte 0 (None) prepended to each scanline of `bytesPerRow` */
function scanlines(
  height: number,
  bytesPerRow: number,
  fill: (row: Uint8Array, y: number) => void,
): Uint8Array {
  const out = new Uint8Array(height * (1 + bytesPerRow))
  for (let y = 0; y < height; y++)
    fill(out.subarray(y * (1 + bytesPerRow) + 1, (y + 1) * (1 + bytesPerRow)), y)
  return out
}

function indexedPng(
  rgba: Uint8Array,
  width: number,
  height: number,
  palette: number[],
): { ihdr: [number, number]; extra: Uint8Array[]; raw: Uint8Array } {
  const index = new Map<number, number>()
  palette.forEach((key, i) => index.set(key, i))
  const bitDepth = palette.length <= 2 ? 1 : palette.length <= 4 ? 2 : palette.length <= 16 ? 4 : 8
  const perByte = 8 / bitDepth
  const bytesPerRow = Math.ceil(width / perByte)
  const raw = scanlines(height, bytesPerRow, (row, y) => {
    const base = y * width * 4
    for (let x = 0; x < width; x++) {
      const i = base + x * 4
      const key =
        ((rgba[i]! << 24) | (rgba[i + 1]! << 16) | (rgba[i + 2]! << 8) | rgba[i + 3]!) >>> 0
      const v = index.get(key)!
      if (bitDepth === 8) row[x] = v
      else
        row[x >> (bitDepth === 1 ? 3 : bitDepth === 2 ? 2 : 1)]! |=
          v << (8 - bitDepth - (x % perByte) * bitDepth)
    }
  })
  const plte = new Uint8Array(palette.length * 3)
  const trns = new Uint8Array(palette.length)
  let lastTransparent = -1
  palette.forEach((key, i) => {
    plte[i * 3] = key >>> 24
    plte[i * 3 + 1] = (key >>> 16) & 0xff
    plte[i * 3 + 2] = (key >>> 8) & 0xff
    trns[i] = key & 0xff
    if (trns[i] !== 255) lastTransparent = i
  })
  const extra = [chunk('PLTE', plte)]
  if (lastTransparent >= 0) extra.push(chunk('tRNS', trns.subarray(0, lastTransparent + 1)))
  return { ihdr: [bitDepth, COLOR_INDEXED], extra, raw }
}

/** Encode an RGBA8 buffer (4 bytes/pixel, tightly packed) as a PNG file. */
export function encodeRgbaPng(rgba: Uint8Array, width: number, height: number): Uint8Array {
  if (rgba.length !== width * height * 4) {
    throw new Error(`rgba buffer is ${rgba.length} bytes, expected ${width * height * 4}`)
  }
  const { opaque, grey, palette } = analyse(rgba)
  let ihdr: [number, number]
  let extra: Uint8Array[] = []
  let raw: Uint8Array
  // opaque greys never exceed 256 values: past the bit-packed depths the
  // palette only adds a PLTE to what greyscale writes directly
  const indexed =
    palette !== undefined && palette.length > 0 && (palette.length <= 16 || !(grey && opaque))
  if (indexed) {
    ;({ ihdr, extra, raw } = indexedPng(rgba, width, height, palette))
  } else if (grey && opaque) {
    ihdr = [8, COLOR_GREY]
    raw = scanlines(height, width, (row, y) => {
      for (let x = 0; x < width; x++) row[x] = rgba[(y * width + x) * 4]!
    })
  } else if (grey) {
    ihdr = [8, COLOR_GREY_ALPHA]
    raw = scanlines(height, width * 2, (row, y) => {
      for (let x = 0; x < width; x++) {
        row[x * 2] = rgba[(y * width + x) * 4]!
        row[x * 2 + 1] = rgba[(y * width + x) * 4 + 3]!
      }
    })
  } else if (opaque) {
    ihdr = [8, COLOR_RGB]
    raw = scanlines(height, width * 3, (row, y) => {
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4
        row[x * 3] = rgba[i]!
        row[x * 3 + 1] = rgba[i + 1]!
        row[x * 3 + 2] = rgba[i + 2]!
      }
    })
  } else {
    ihdr = [8, COLOR_RGBA]
    raw = scanlines(height, width * 4, (row, y) => {
      row.set(rgba.subarray(y * width * 4, (y + 1) * width * 4))
    })
  }

  const header = new Uint8Array(13)
  writeU32(header, 0, width)
  writeU32(header, 4, height)
  header[8] = ihdr[0]
  header[9] = ihdr[1]
  // compression/filter/interlace all 0

  const signature = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  const parts = [
    signature,
    chunk('IHDR', header),
    ...extra,
    chunk('IDAT', new Uint8Array(deflateSync(raw, { level: DEFLATE_LEVEL }))),
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
