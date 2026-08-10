/**
 * TIFF → PNG transcoding. Chromium cannot decode image/tiff, so pictures
 * embedded as word/media/*.tif(f) would render as blank frames. Decode with
 * UTIF (pure JS) and re-encode as PNG via canvas for display; the original
 * TIFF bytes stay untouched in the package so saving preserves them.
 */
import UTIF from 'utif2'

const TIFF_MIMES = new Set(['image/tiff', 'image/tif', 'image/x-tiff'])

export function isTiffMime(mime: string | undefined): mime is string {
  return mime !== undefined && TIFF_MIMES.has(mime)
}

// the engine compiles without DOM libs (it also runs under node); the canvas
// re-encode is reached only in renderer environments, typed structurally here
interface CanvasLike {
  width: number
  height: number
  getContext(id: '2d'): { putImageData(data: unknown, x: number, y: number): void } | null
  toDataURL(type: string): string
}
interface DomGlobals {
  document?: { createElement(tag: 'canvas'): CanvasLike }
  ImageData?: new (data: Uint8ClampedArray, w: number, h: number) => unknown
}

/**
 * Render TIFF bytes to a PNG data URL. Returns null on parse failure or when
 * no canvas API exists (non-renderer environments), so callers keep their
 * existing empty-frame degrade.
 */
export function tiffToDataUrl(bytes: ArrayBuffer | Uint8Array): string | null {
  const dom = globalThis as DomGlobals
  if (!dom.document || !dom.ImageData) return null
  try {
    // copy into a fresh ArrayBuffer (a Uint8Array view may sit on a
    // SharedArrayBuffer, which UTIF's signature rejects)
    const buf = bytes instanceof Uint8Array ? new Uint8Array(bytes).buffer : bytes
    const ifds = UTIF.decode(buf)
    if (!ifds.length) return null
    // multi-page/multi-resolution TIFFs: pick the largest page
    let page = ifds[0]
    for (const ifd of ifds) {
      UTIF.decodeImage(buf, ifd)
      if ((ifd.width || 0) * (ifd.height || 0) > (page.width || 0) * (page.height || 0)) page = ifd
    }
    const width = page.width
    const height = page.height
    if (!width || !height) return null
    const rgba = UTIF.toRGBA8(page)
    const canvas = dom.document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx2d = canvas.getContext('2d')
    if (!ctx2d) return null
    const pixels = new Uint8ClampedArray(rgba.buffer, rgba.byteOffset, width * height * 4)
    ctx2d.putImageData(new dom.ImageData(pixels, width, height), 0, 0)
    return canvas.toDataURL('image/png')
  } catch {
    return null
  }
}
