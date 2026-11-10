/**
 * TIFF → PNG transcoding (main process). Chromium cannot decode TIFF, so pictures
 * embedded as ppt/media/*.tif(f) would render as blank placeholders. Decode with
 * UTIF (pure JS) and re-encode as PNG for display; the original TIFF bytes stay
 * untouched in the package so saving preserves them byte-for-byte.
 */
import UTIF from 'utif2'
import { PNG } from 'pngjs'

export interface DecodedTiff {
  png: Uint8Array
  width: number
  height: number
}

/** Pixel budget: a crafted ppt/media/*.tif with giant IFD dims would OOM
 *  the main process in toRGBA8 (w*h*4 bytes). Fail closed (blank placeholder)
 *  past 256MP or 32767px per side (a 600 dpi A3 scan is ~69MP). */
const MAX_TIFF_PIXELS = 256 * 1024 * 1024
const MAX_TIFF_DIM = 32767

function tiffDimsOk(width: unknown, height: unknown): width is number {
  return (
    typeof width === 'number' &&
    typeof height === 'number' &&
    Number.isInteger(width) &&
    Number.isInteger(height) &&
    width > 0 &&
    height > 0 &&
    width <= MAX_TIFF_DIM &&
    height <= MAX_TIFF_DIM &&
    width * height <= MAX_TIFF_PIXELS
  )
}

export function tiffToPng(bytes: Uint8Array): DecodedTiff | null {
  try {
    const buf = Buffer.from(bytes)
    const ifds = UTIF.decode(buf)
    if (!ifds.length) return null
    // Multi-page/multi-resolution TIFFs: pick the largest page within budget.
    // UTIF only fills .width/.height during decodeImage (which allocates the
    // pixels), so read the raw header tags (t256/t257) for the budget check
    // and decodeImage solely the chosen page.
    const headerDims = (ifd: Record<string, unknown>): { width: unknown; height: unknown } => ({
      width: Array.isArray(ifd.t256) ? ifd.t256[0] : (ifd as { width?: unknown }).width,
      height: Array.isArray(ifd.t257) ? ifd.t257[0] : (ifd as { height?: unknown }).height,
    })
    let page: (typeof ifds)[number] | null = null
    let pagePixels = 0
    for (const ifd of ifds) {
      const dims = headerDims(ifd)
      if (!tiffDimsOk(dims.width, dims.height)) continue
      const cur = dims.width * (dims.height as number)
      if (!page || cur > pagePixels) {
        page = ifd
        pagePixels = cur
      }
    }
    if (!page) return null
    UTIF.decodeImage(buf, page)
    const width = page.width as number
    const height = page.height as number
    if (!tiffDimsOk(width, height)) return null
    const rgba = UTIF.toRGBA8(page)
    const png = new PNG({ width, height })
    png.data = Buffer.from(rgba.buffer, rgba.byteOffset, rgba.byteLength)
    return { png: PNG.sync.write(png), width, height }
  } catch {
    return null
  }
}
