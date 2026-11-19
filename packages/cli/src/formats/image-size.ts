/** Pixel size from the header of a PNG, JPEG or GIF; null for anything else. */
/** Largest dimension honored: beyond this the header is corrupt. */
const MAX_IMAGE_DIM = 100_000

function validDims(width: number, height: number): boolean {
  return (
    Number.isInteger(width) &&
    Number.isInteger(height) &&
    width > 0 &&
    height > 0 &&
    width <= MAX_IMAGE_DIM &&
    height <= MAX_IMAGE_DIM
  )
}

export function imageSize(
  bytes: Uint8Array,
): { width: number; height: number; mime: 'image/png' | 'image/jpeg' | 'image/gif' } | null {
  const b = bytes
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    const width = readU32(b, 16)
    const height = readU32(b, 20)
    return validDims(width, height) ? { width, height, mime: 'image/png' } : null
  }
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
    const width = b[6]! | (b[7]! << 8)
    const height = b[8]! | (b[9]! << 8)
    return validDims(width, height) ? { width, height, mime: 'image/gif' } : null
  }
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) return null
      const marker = b[i + 1]!
      // Standalone markers carry no length word (RSTn, SOI, EOI, TEM):
      // skipping blindly would land mid-frame and misparse the dims.
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9)) {
        i += 2
        continue
      }
      const size = (b[i + 2]! << 8) | b[i + 3]!
      if (size < 2) return null
      // SOF0..SOF15 except DHT (C4), JPG (C8), DAC (CC) carry the frame size
      if (
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc
      ) {
        const height = (b[i + 5]! << 8) | b[i + 6]!
        const width = (b[i + 7]! << 8) | b[i + 8]!
        return validDims(width, height) ? { width, height, mime: 'image/jpeg' } : null
      }
      i += 2 + size
    }
  }
  return null
}

function readU32(b: Uint8Array, at: number): number {
  return ((b[at]! << 24) | (b[at + 1]! << 16) | (b[at + 2]! << 8) | b[at + 3]!) >>> 0
}
