/**
 * Pure pixel/geometry helpers for baked image edits (flip / transparency / crop).
 * PDF content-stream images are plain bitmaps, so these edits rewrite pixels and land
 * as replaceImage ops; no DOM here — App owns the canvas decode/encode glue.
 */
import type { PixelImage } from './cutout'

type Rect = readonly [number, number, number, number]

/** Kept region of a crop, as 0..1 fractions of the displayed image (l<r, t<b) */
export interface CropFractions {
  l: number
  t: number
  r: number
  b: number
}

/** Mirror pixels horizontally ('h') or vertically ('v'); returns a new array */
export function flipPixels(img: PixelImage, axis: 'h' | 'v'): Uint8ClampedArray<ArrayBuffer> {
  const { data, width: w, height: h } = img
  const out = new Uint8ClampedArray(data.length)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const src = (y * w + x) * 4
      const dst = axis === 'h' ? (y * w + (w - 1 - x)) * 4 : ((h - 1 - y) * w + x) * 4
      out[dst] = data[src]!
      out[dst + 1] = data[src + 1]!
      out[dst + 2] = data[src + 2]!
      out[dst + 3] = data[src + 3]!
    }
  }
  return out
}

/** Scale every pixel's alpha by factor (0..1); returns a new array */
export function multiplyAlpha(img: PixelImage, factor: number): Uint8ClampedArray<ArrayBuffer> {
  const f = Math.min(1, Math.max(0, factor))
  const out = new Uint8ClampedArray(img.data)
  for (let i = 3; i < out.length; i += 4) out[i] = Math.round(out[i]! * f)
  return out
}

/**
 * PDF user-space footprint of the kept crop region. Fractions are in the image's
 * displayed orientation (top-left origin), so the display top edge maps to y2 (PDF y-up).
 * Object-space math — exact for existing images, whose bounds-based ops treat the
 * bitmap as axis-aligned in page space.
 */
export function cropRect(rect: Rect, crop: CropFractions): [number, number, number, number] {
  const w = rect[2] - rect[0]
  const h = rect[3] - rect[1]
  return [rect[0] + crop.l * w, rect[3] - crop.b * h, rect[0] + crop.r * w, rect[3] - crop.t * h]
}
