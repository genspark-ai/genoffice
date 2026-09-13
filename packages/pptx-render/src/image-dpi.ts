/**
 * Image resolution tag (PNG pHYs / JPEG JFIF density) read from a base64 data URL.
 * PowerPoint sizes a:tile cells by this tag (measured: a 128px 75dpi JFIF texture tiles
 * at 1.71in, a 1460px 150dpi PNG at 9.73in); untagged bitmaps fall back to 144dpi.
 */

/** Only the header region is decoded: pHYs precedes IDAT and JFIF's APP0 is at offset 2. */
const HEAD_B64_CHARS = 96 * 1024

const cache = new Map<string, { x: number; y: number } | undefined>()

export function imageDpiFromDataUrl(
  dataUrl: string | undefined,
): { x: number; y: number } | undefined {
  if (!dataUrl) return undefined
  if (cache.has(dataUrl)) return cache.get(dataUrl)
  const dpi = readDpi(decodeHead(dataUrl))
  if (cache.size > 256) cache.clear()
  cache.set(dataUrl, dpi)
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

function readDpi(b: Uint8Array): { x: number; y: number } | undefined {
  if (b.length < 16) return undefined
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return pngDpi(b)
  if (b[0] === 0xff && b[1] === 0xd8) return jpegDpi(b)
  return undefined
}

function plausible(x: number, y: number): { x: number; y: number } | undefined {
  return x >= 10 && x <= 10000 && y >= 10 && y <= 10000 ? { x, y } : undefined
}

function u32(b: Uint8Array, o: number): number {
  return ((b[o]! << 24) >>> 0) + (b[o + 1]! << 16) + (b[o + 2]! << 8) + b[o + 3]!
}

function pngDpi(b: Uint8Array): { x: number; y: number } | undefined {
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

function jpegDpi(b: Uint8Array): { x: number; y: number } | undefined {
  let o = 2
  while (o + 4 <= b.length && b[o] === 0xff) {
    const marker = b[o + 1]!
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
      o += 2
      continue
    }
    const len = (b[o + 2]! << 8) + b[o + 3]!
    if (marker === 0xda || marker === 0xd9) return undefined
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
        if (units === 1) return plausible(xd, yd)
        if (units === 2) return plausible(xd * 2.54, yd * 2.54)
        return undefined
      }
    }
    o += 2 + len
  }
  return undefined
}
