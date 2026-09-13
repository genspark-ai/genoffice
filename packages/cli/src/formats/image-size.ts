/** Pixel size from the header of a PNG, JPEG or GIF; null for anything else. */
export function imageSize(
  bytes: Uint8Array,
): { width: number; height: number; mime: 'image/png' | 'image/jpeg' | 'image/gif' } | null {
  const b = bytes
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    return { width: readU32(b, 16), height: readU32(b, 20), mime: 'image/png' }
  }
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
    return { width: b[6]! | (b[7]! << 8), height: b[8]! | (b[9]! << 8), mime: 'image/gif' }
  }
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) return null
      const marker = b[i + 1]!
      const size = (b[i + 2]! << 8) | b[i + 3]!
      // SOF0..SOF15 except DHT (C4), JPG (C8), DAC (CC) carry the frame size
      if (
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc
      ) {
        return {
          height: (b[i + 5]! << 8) | b[i + 6]!,
          width: (b[i + 7]! << 8) | b[i + 8]!,
          mime: 'image/jpeg',
        }
      }
      i += 2 + size
    }
  }
  return null
}

function readU32(b: Uint8Array, at: number): number {
  return ((b[at]! << 24) | (b[at + 1]! << 16) | (b[at + 2]! << 8) | b[at + 3]!) >>> 0
}
