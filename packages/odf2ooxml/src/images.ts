/**
 * Intrinsic pixel sizes of the image formats ODF embeds (png/jpeg/gif), read
 * straight from the byte stream — no image library. Used when a draw:frame
 * declares no svg:width/height and the import must still land the picture at
 * its natural size.
 */

export function intrinsicImageSize(
  bytes: Uint8Array,
  mime: 'image/png' | 'image/jpeg' | 'image/gif',
): { width: number; height: number } | null {
  if (mime === 'image/png') return pngSize(bytes)
  if (mime === 'image/gif') return gifSize(bytes)
  return jpegSize(bytes)
}

function pngSize(bytes: Uint8Array): { width: number; height: number } | null {
  // fixed 8-byte signature then the IHDR chunk: length(4) type(4) width(4) height(4)
  if (bytes.length < 24) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (view.getUint32(8) !== 13) return null
  if (bytes[12] !== 0x49 || bytes[13] !== 0x48 || bytes[14] !== 0x44 || bytes[15] !== 0x52)
    return null
  return { width: view.getUint32(16), height: view.getUint32(20) }
}

function gifSize(bytes: Uint8Array): { width: number; height: number } | null {
  if (bytes.length < 10) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return { width: view.getUint16(6, true), height: view.getUint16(8, true) }
}

/** walk the segment chain to the first SOFn frame header */
function jpegSize(bytes: Uint8Array): { width: number; height: number } | null {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return null
  let offset = 2
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset++
      continue
    }
    const marker = bytes[offset + 1]
    // standalone markers carry no length segment
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd9) || marker === 0x01) {
      offset += 2
      continue
    }
    const length = (bytes[offset + 2] << 8) | bytes[offset + 3]
    // SOF0-SOF15 except DHT (0xc4), JPG (0xc8), DAC (0xcc) carry the dimensions
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      if (offset + 9 > bytes.length) return null
      return {
        height: (bytes[offset + 5] << 8) | bytes[offset + 6],
        width: (bytes[offset + 7] << 8) | bytes[offset + 8],
      }
    }
    offset += 2 + length
  }
  return null
}
