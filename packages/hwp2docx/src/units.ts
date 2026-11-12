/**
 * HWPUNIT is 1/7200 inch (Hangul's internal unit; 1pt = 100 HWPUNIT).
 * OOXML twips are 1/1440 inch, EMU 1/914400 inch, CSS px 1/96 inch.
 */

export function hwpToTwips(value: number): number {
  return Math.round(value / 5)
}

export function hwpToPx(value: number): number {
  return Math.max(1, Math.round(value / 75))
}

export function hwpToEmu(value: number): number {
  return Math.round(value * 127)
}

/** sniff the real image type from magic bytes (BinData extensions lie sometimes) */
export function sniffImage(
  bytes: Uint8Array,
): 'png' | 'jpeg' | 'gif' | 'bmp' | 'wmf' | 'emf' | null {
  const b = bytes
  if (b.length < 4) return null
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png'
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpeg'
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif'
  if (b[0] === 0x42 && b[1] === 0x4d) return 'bmp'
  if (b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) return 'wmf'
  if (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x00 && b[3] === 0x00 && b.length > 44) {
    if (b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46) return 'emf'
  }
  return null
}

/** A producer-controlled image header must not hand zero, NaN or a 32-bit
    maximum straight to an OOXML drawing extent (cx/cy are written verbatim).
    Keep the same 40M pixel budget bmp.ts uses for embedded bitmaps; the caller
    already falls back to a default size when intrinsicSize returns null. */
const MAX_IMAGE_PIXELS = 40_000_000

function saneSize(width: number, height: number): { width: number; height: number } | null {
  if (
    !Number.isInteger(width) ||
    !Number.isInteger(height) ||
    width < 1 ||
    height < 1 ||
    width > 0x7fffffff ||
    height > 0x7fffffff ||
    width * height > MAX_IMAGE_PIXELS
  )
    return null
  return { width, height }
}

/** intrinsic pixel size for png / jpeg / gif headers */
export function intrinsicSize(bytes: Uint8Array): { width: number; height: number } | null {
  const kind = sniffImage(bytes)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (kind === 'png' && bytes.length >= 24) {
    return saneSize(view.getUint32(16), view.getUint32(20))
  }
  if (kind === 'gif' && bytes.length >= 10) {
    return saneSize(view.getUint16(6, true), view.getUint16(8, true))
  }
  if (kind === 'jpeg') {
    let offset = 2
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) return null
      const marker = bytes[offset + 1]
      const length = view.getUint16(offset + 2)
      if (
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc
      ) {
        return saneSize(view.getUint16(offset + 7), view.getUint16(offset + 5))
      }
      offset += 2 + length
    }
  }
  return null
}
