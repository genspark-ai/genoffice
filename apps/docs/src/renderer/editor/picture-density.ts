/**
 * The resolution a picture file states for itself. Word sizes a picture from
 * it: a 600 px wide picture at 300 dpi is two inches wide, not 6.25.
 */
export interface PictureDensity {
  xDpi: number
  yDpi: number
}

export const SCREEN_DPI = 96

const METRES_PER_INCH = 0.0254
const CM_PER_INCH = 2.54

// 11811 px/m is how a file says 300 dpi
const whole = (dpi: number) => (Math.abs(dpi - Math.round(dpi)) < 0.05 ? Math.round(dpi) : dpi)

function density(x: number, y: number): PictureDensity | null {
  return Number.isFinite(x) && Number.isFinite(y) && x > 0 && y > 0
    ? { xDpi: whole(x), yDpi: whole(y) }
    : null
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]

function pngDensity(bytes: Uint8Array, view: DataView): PictureDensity | null {
  let at = PNG_SIGNATURE.length
  while (at + 12 <= bytes.length) {
    const length = view.getUint32(at)
    const type = String.fromCharCode(...bytes.subarray(at + 4, at + 8))
    if (type === 'IDAT' || type === 'IEND') return null
    if (type === 'pHYs' && length === 9 && at + 17 <= bytes.length) {
      if (bytes[at + 16] !== 1) return null
      return density(
        view.getUint32(at + 8) * METRES_PER_INCH,
        view.getUint32(at + 12) * METRES_PER_INCH,
      )
    }
    at += 12 + length
  }
  return null
}

function rational(view: DataView, at: number, little: boolean): number {
  if (at + 8 > view.byteLength) return NaN
  return view.getUint32(at, little) / view.getUint32(at + 4, little)
}

/** XResolution / YResolution / ResolutionUnit of the TIFF header an Exif segment holds */
function exifDensity(view: DataView, tiff: number, end: number): PictureDensity | null {
  if (tiff + 8 > end) return null
  const order = view.getUint16(tiff)
  if (order !== 0x4949 && order !== 0x4d4d) return null
  const little = order === 0x4949
  const ifd = tiff + view.getUint32(tiff + 4, little)
  if (ifd + 2 > end) return null
  const count = view.getUint16(ifd, little)
  let x = NaN
  let y = NaN
  let unit = 2
  for (let i = 0; i < count; i++) {
    const entry = ifd + 2 + i * 12
    if (entry + 12 > end) break
    const tag = view.getUint16(entry, little)
    if (tag === 0x011a) x = rational(view, tiff + view.getUint32(entry + 8, little), little)
    else if (tag === 0x011b) y = rational(view, tiff + view.getUint32(entry + 8, little), little)
    else if (tag === 0x0128) unit = view.getUint16(entry + 8, little)
  }
  if (unit === 3) return density(x * CM_PER_INCH, y * CM_PER_INCH)
  return unit === 2 ? density(x, y) : null
}

const ascii = (bytes: Uint8Array, at: number, text: string) =>
  [...text].every((ch, i) => bytes[at + i] === ch.charCodeAt(0))

/** a JFIF header that gives a unit wins; otherwise the Exif resolution */
function jpegDensity(bytes: Uint8Array, view: DataView): PictureDensity | null {
  let exif: PictureDensity | null = null
  let at = 2
  while (at + 4 <= bytes.length && bytes[at] === 0xff) {
    const marker = bytes[at + 1]
    if (marker === 0xda || marker === 0xd9) break
    const length = view.getUint16(at + 2)
    const body = at + 4
    const end = Math.min(bytes.length, at + 2 + length)
    if (marker === 0xe0 && ascii(bytes, body, 'JFIF\0') && body + 12 <= end) {
      const unit = bytes[body + 7]
      const x = view.getUint16(body + 8)
      const y = view.getUint16(body + 10)
      const stated =
        unit === 1 ? density(x, y) : unit === 2 ? density(x * CM_PER_INCH, y * CM_PER_INCH) : null
      if (stated) return stated
    } else if (marker === 0xe1 && ascii(bytes, body, 'Exif\0\0')) {
      exif ??= exifDensity(view, body + 6, end)
    }
    at += 2 + length
  }
  return exif
}

const BMP_FILE_HEADER = 14
const BMP_INFO_HEADER = 40

/** biXPelsPerMeter / biYPelsPerMeter of a BITMAPINFOHEADER or later; the older core header has none */
function bmpDensity(bytes: Uint8Array, view: DataView): PictureDensity | null {
  if (bytes.length < BMP_FILE_HEADER + BMP_INFO_HEADER) return null
  if (view.getUint32(BMP_FILE_HEADER, true) < BMP_INFO_HEADER) return null
  return density(
    view.getInt32(BMP_FILE_HEADER + 24, true) * METRES_PER_INCH,
    view.getInt32(BMP_FILE_HEADER + 28, true) * METRES_PER_INCH,
  )
}

/**
 * The density a PNG (pHYs), JPEG (JFIF, Exif) or BMP (pels per metre) states;
 * null when it states none. A GIF has no field for one.
 */
export function pictureDensityOf(bytes: Uint8Array): PictureDensity | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  try {
    if (PNG_SIGNATURE.every((byte, i) => bytes[i] === byte)) return pngDensity(bytes, view)
    if (bytes[0] === 0xff && bytes[1] === 0xd8) return jpegDensity(bytes, view)
    if (bytes[0] === 0x42 && bytes[1] === 0x4d) return bmpDensity(bytes, view)
  } catch {
    return null
  }
  return null
}

/** the size Word calls original, in CSS px: the pixels at the file's density, 96 dpi when it states none */
export function originalPictureSize(
  pixels: { width: number; height: number },
  density: PictureDensity | null,
): { width: number; height: number } {
  if (!density) return pixels
  return {
    width: (pixels.width * SCREEN_DPI) / density.xDpi,
    height: (pixels.height * SCREEN_DPI) / density.yDpi,
  }
}
