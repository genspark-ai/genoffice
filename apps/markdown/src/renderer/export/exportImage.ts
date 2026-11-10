import type { NewImage } from '@genoffice/docx-engine'
import type { ImageData } from '../../shared/ipc'
import { exportImageMime } from '../../shared/export-image-mime'
import { parseSvgViewBox, pinSvgIntrinsicSize } from '../editor/diagrams'

type NativeMime = NewImage['mime']

const NATIVE_MIMES: ReadonlySet<string> = new Set<NativeMime>([
  'image/png',
  'image/jpeg',
  'image/gif',
])

/** Doc size when the picture cannot be decoded at all */
export const FALLBACK_IMAGE_SIZE = { width: 400, height: 300 }
/** Hard cap for a raster dimension so a huge SVG cannot allocate an unbounded bitmap */
export const MAX_RASTER_DIM_PX = 4096
/** HiDPI scale for vector sources; bitmaps never upscale */
const VECTOR_RASTER_SCALE = 2

export interface ImageSize {
  width: number
  height: number
}

export interface DecodedImage extends ImageSize {
  source: CanvasImageSource
}

export type ImageDecoder = (dataUrl: string) => Promise<DecodedImage | null>

/** `![x](data:…)` needs no main-process round trip */
export function decodeImageDataUrl(src: string): ImageData | null {
  const match = /^data:([^;,]+)((?:;[^;,]*)*),(.*)$/is.exec(src.trim())
  if (!match) return null
  const mime = exportImageMime(match[1])
  if (!mime) return null
  const payload = match[3]!
  try {
    if (/;base64/i.test(match[2]!)) {
      const base64 = payload.replace(/\s+/g, '')
      atob(base64)
      return { base64, mime }
    }
    return { base64: bytesToBase64(new TextEncoder().encode(percentDecodeLenient(payload))), mime }
  } catch {
    return null
  }
}

/** Browsers keep a bare `%` (SVG `width="100%"`) literal; decodeURIComponent would throw */
function percentDecodeLenient(payload: string): string {
  return decodeURIComponent(payload.replace(/%(?![0-9a-f]{2})/gi, '%25'))
}

export function fitDocxWidth(
  size: ImageSize,
  maxWidthPx: number,
): { widthPx: number; heightPx: number } {
  const width = Math.max(1, Math.round(size.width))
  const height = Math.max(1, Math.round(size.height))
  if (width <= maxWidthPx) return { widthPx: width, heightPx: height }
  return { widthPx: maxWidthPx, heightPx: Math.max(1, Math.round((height * maxWidthPx) / width)) }
}

/** Backing bitmap size: vectors draw at HiDPI scale, bitmaps at most at their own size */
export function rasterSize(
  doc: { widthPx: number; heightPx: number },
  intrinsic: ImageSize,
  vector: boolean,
): ImageSize {
  let width = vector
    ? doc.widthPx * VECTOR_RASTER_SCALE
    : Math.min(intrinsic.width, doc.widthPx * VECTOR_RASTER_SCALE)
  let height = Math.round((doc.heightPx * width) / doc.widthPx)
  const peak = Math.max(width, height)
  if (peak > MAX_RASTER_DIM_PX) {
    width = Math.floor((width * MAX_RASTER_DIM_PX) / peak)
    height = Math.floor((height * MAX_RASTER_DIM_PX) / peak)
  }
  return { width: Math.max(1, Math.round(width)), height: Math.max(1, height) }
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  }
  return btoa(binary)
}

function base64ToText(base64: string): string {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return new TextDecoder().decode(bytes)
}

const decodeWithImage: ImageDecoder = (dataUrl) =>
  new Promise((resolve) => {
    const img = new Image()
    img.onload = () => {
      resolve(
        img.naturalWidth > 0 && img.naturalHeight > 0
          ? { width: img.naturalWidth, height: img.naturalHeight, source: img }
          : null,
      )
    }
    img.onerror = () => resolve(null)
    img.src = dataUrl
  })

/**
 * SVG bytes with a usable intrinsic size: `<img>` decodes a sizeless root at
 * 300x150, so the viewBox is pinned onto the root first.
 */
function sizedSvgDataUrl(base64: string): { dataUrl: string; viewBox: ImageSize | null } {
  const svg = base64ToText(base64)
  const viewBox = parseSvgViewBox(svg)
  const sized = viewBox ? pinSvgIntrinsicSize(svg, viewBox.width, viewBox.height) : svg
  return {
    dataUrl: `data:image/svg+xml;base64,${bytesToBase64(new TextEncoder().encode(sized))}`,
    viewBox,
  }
}

/**
 * Picture bytes for the DOCX: PNG/JPEG/GIF land as they are, every other
 * displayable format (SVG, WebP, BMP, AVIF) is drawn onto a canvas and
 * re-encoded as PNG, since that is all the docx-engine save path accepts.
 * Null when the bytes cannot be decoded; the caller falls back to alt text.
 */
export async function toDocxImage(
  data: ImageData,
  maxWidthPx: number,
  decode: ImageDecoder = decodeWithImage,
): Promise<NewImage | null> {
  if (NATIVE_MIMES.has(data.mime)) {
    const size = (await decode(`data:${data.mime};base64,${data.base64}`)) ?? FALLBACK_IMAGE_SIZE
    return { base64: data.base64, mime: data.mime as NativeMime, ...fitDocxWidth(size, maxWidthPx) }
  }
  const vector = data.mime === 'image/svg+xml'
  let dataUrl = `data:${data.mime};base64,${data.base64}`
  let viewBox: ImageSize | null = null
  if (vector) {
    try {
      ;({ dataUrl, viewBox } = sizedSvgDataUrl(data.base64))
    } catch {
      return null
    }
  }
  const decoded = await decode(dataUrl)
  if (!decoded) return null
  const intrinsic: ImageSize = vector && viewBox ? viewBox : decoded
  const doc = fitDocxWidth(intrinsic, maxWidthPx)
  const raster = rasterSize(doc, intrinsic, vector)
  const canvas = document.createElement('canvas')
  canvas.width = raster.width
  canvas.height = raster.height
  const ctx = canvas.getContext('2d')
  if (!ctx) return null
  try {
    ctx.drawImage(decoded.source, 0, 0, raster.width, raster.height)
    const png = canvas.toDataURL('image/png')
    const base64 = png.slice(png.indexOf(',') + 1)
    return base64 ? { base64, mime: 'image/png', ...doc } : null
  } catch {
    return null
  } finally {
    canvas.width = canvas.height = 0
  }
}
