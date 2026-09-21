/**
 * Image resolution tag (PNG pHYs / JPEG JFIF density, EXIF X/YResolution when JFIF
 * declares none) read from a base64 data URL or raw bytes.
 * PowerPoint sizes a:tile cells by this tag (measured: a 128px 75dpi JFIF texture tiles
 * at 1.71in, a 1460px 150dpi PNG at 9.73in); untagged bitmaps fall back to 144dpi.
 */

export interface ImageDpi {
  x: number
  y: number
}

/** Only the header region is decoded: pHYs precedes IDAT and JFIF's APP0 is at offset 2. */
const HEAD_B64_CHARS = 96 * 1024

const cache = new Map<string, ImageDpi | undefined>()

/** Max cached entries; oldest inserted key is evicted once the cap is reached. */
export const IMAGE_DPI_CACHE_MAX = 256

/**
 * Compact cache key for a data URL: input length plus FNV-1a hash of the head.
 * The dpi tag lives in the header decodeHead reads, so hashing only that region
 * keeps the key exact while bounding the cost for multi-megabyte URLs.
 */
export function cacheKeyFor(dataUrl: string): string {
  const comma = dataUrl.indexOf(',')
  const end = Math.min(dataUrl.length, (comma < 0 ? 0 : comma + 1) + HEAD_B64_CHARS)
  let hash = 0x811c9dc5
  for (let i = 0; i < end; i++) {
    hash ^= dataUrl.charCodeAt(i)
    hash = Math.imul(hash, 0x01000193)
  }
  return `${dataUrl.length}:${(hash >>> 0).toString(16).padStart(8, '0')}`
}

/** Current number of cached entries (test helper). */
export function imageDpiCacheSize(): number {
  return cache.size
}

/** Current cache keys in insertion order (test helper). */
export function imageDpiCacheKeys(): string[] {
  return [...cache.keys()]
}

/** Empty the cache (test helper). */
export function clearImageDpiCache(): void {
  cache.clear()
}

export function imageDpiFromDataUrl(dataUrl: string | undefined): ImageDpi | undefined {
  if (!dataUrl) return undefined
  const key = cacheKeyFor(dataUrl)
  if (cache.has(key)) {
    // Refresh recency so eviction drops the least recently used entry.
    const cached = cache.get(key)
    cache.delete(key)
    cache.set(key, cached)
    return cached
  }
  const dpi = imageDpiFromBytes(decodeHead(dataUrl))
  if (cache.size >= IMAGE_DPI_CACHE_MAX) {
    const oldest = cache.keys().next()
    if (!oldest.done) cache.delete(oldest.value)
  }
  cache.set(key, dpi)
  return dpi
}

function decodeHead(dataUrl: string): Uint8Array {
  const comma = dataUrl.indexOf(',')
  if (comma < 0 || !/;base64/i.test(dataUrl.slice(0, comma))) return new Uint8Array(0)
  const b64 = dataUrl.slice(comma + 1, comma + 1 + HEAD_B64_CHARS)
  const whole = Math.floor(b64.length / 4) * 4
  const chunk = b64.length < HEAD_B64_CHARS ? b64 : b64.slice(0, whole)
  try {
    if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(chunk, 'base64'))
    const s = atob(chunk)
    const out = new Uint8Array(s.length)
    for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i)
    return out
  } catch {
    return new Uint8Array(0)
  }
}

/** Declared resolution of a PNG or JPEG; undefined when the header carries none. */
export function imageDpiFromBytes(b: Uint8Array): ImageDpi | undefined {
  if (b.length < 16) return undefined
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return pngDpi(b)
  if (b[0] === 0xff && b[1] === 0xd8) return jpegDpi(b)
  return undefined
}

function plausible(x: number, y: number): ImageDpi | undefined {
  return x >= 10 && x <= 10000 && y >= 10 && y <= 10000 ? { x, y } : undefined
}

function u32(b: Uint8Array, o: number): number {
  return ((b[o]! << 24) >>> 0) + (b[o + 1]! << 16) + (b[o + 2]! << 8) + b[o + 3]!
}

function pngDpi(b: Uint8Array): ImageDpi | undefined {
  let o = 8
  while (o + 12 <= b.length) {
    const len = u32(b, o)
    const type = String.fromCharCode(b[o + 4]!, b[o + 5]!, b[o + 6]!, b[o + 7]!)
    if (type === 'IDAT' || type === 'IEND') return undefined
    if (type === 'pHYs' && len >= 9 && o + 8 + 9 <= b.length) {
      // unit 1 = pixels per metre; unit 0 only states an aspect ratio
      if (b[o + 16] !== 1) return undefined
      return plausible(u32(b, o + 8) * 0.0254, u32(b, o + 12) * 0.0254)
    }
    o += 12 + len
  }
  return undefined
}

function jpegDpi(b: Uint8Array): ImageDpi | undefined {
  let exif: ImageDpi | undefined
  let o = 2
  while (o + 4 <= b.length && b[o] === 0xff) {
    const marker = b[o + 1]!
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
      o += 2
      continue
    }
    const len = (b[o + 2]! << 8) + b[o + 3]!
    if (marker === 0xda || marker === 0xd9) return exif
    if (marker === 0xe0 && len >= 16 && o + 2 + len <= b.length) {
      const p = o + 4
      const isJfif =
        b[p] === 0x4a &&
        b[p + 1] === 0x46 &&
        b[p + 2] === 0x49 &&
        b[p + 3] === 0x46 &&
        b[p + 4] === 0
      if (isJfif) {
        const units = b[p + 7]!
        const xd = (b[p + 8]! << 8) + b[p + 9]!
        const yd = (b[p + 10]! << 8) + b[p + 11]!
        // JFIF wins over EXIF; units 0 declares only an aspect ratio, so keep scanning
        if (units === 1) return plausible(xd, yd)
        if (units === 2) return plausible(xd * 2.54, yd * 2.54)
      }
    }
    if (marker === 0xe1 && !exif && o + 2 + len <= b.length) exif = exifDpi(b, o + 4, o + 2 + len)
    o += 2 + len
  }
  return exif
}

const EXIF_X_RESOLUTION = 0x011a
const EXIF_Y_RESOLUTION = 0x011b
const EXIF_RESOLUTION_UNIT = 0x0128

/** IFD0 XResolution/YResolution of an APP1 "Exif" segment spanning [p, end). */
function exifDpi(b: Uint8Array, p: number, end: number): ImageDpi | undefined {
  const isExif =
    b[p] === 0x45 &&
    b[p + 1] === 0x78 &&
    b[p + 2] === 0x69 &&
    b[p + 3] === 0x66 &&
    b[p + 4] === 0 &&
    b[p + 5] === 0
  if (!isExif) return undefined
  const tiff = p + 6
  const le = b[tiff] === 0x49 && b[tiff + 1] === 0x49
  if (!le && !(b[tiff] === 0x4d && b[tiff + 1] === 0x4d)) return undefined
  const u16 = (o: number) => (le ? b[o]! | (b[o + 1]! << 8) : (b[o]! << 8) | b[o + 1]!)
  const rd32 = (o: number) =>
    le ? (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0 : u32(b, o)
  const rational = (e: number) => {
    const at = tiff + rd32(e + 8)
    if (at + 8 > end) return undefined
    const den = rd32(at + 4)
    return den ? rd32(at) / den : undefined
  }
  if (tiff + 8 > end) return undefined
  const ifd0 = tiff + rd32(tiff + 4)
  if (ifd0 + 2 > end) return undefined
  const count = u16(ifd0)
  let x: number | undefined
  let y: number | undefined
  let unit = 2
  for (let i = 0; i < count; i++) {
    const e = ifd0 + 2 + i * 12
    if (e + 12 > end) return undefined
    const tag = u16(e)
    if (tag === EXIF_X_RESOLUTION) x = rational(e)
    else if (tag === EXIF_Y_RESOLUTION) y = rational(e)
    else if (tag === EXIF_RESOLUTION_UNIT) unit = u16(e + 8)
  }
  if (x == null || y == null) return undefined
  if (unit === 3) return plausible(x * 2.54, y * 2.54)
  return unit === 2 ? plausible(x, y) : undefined
}
