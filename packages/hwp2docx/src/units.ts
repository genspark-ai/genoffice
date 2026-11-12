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

/** intrinsic pixel size for png / jpeg / gif headers */
export function intrinsicSize(bytes: Uint8Array): { width: number; height: number } | null {
  const kind = sniffImage(bytes)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (kind === 'png' && bytes.length >= 24) {
    return { width: view.getUint32(16), height: view.getUint32(20) }
  }
  if (kind === 'gif' && bytes.length >= 10) {
    return { width: view.getUint16(6, true), height: view.getUint16(8, true) }
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
        return { width: view.getUint16(offset + 7), height: view.getUint16(offset + 5) }
      }
      offset += 2 + length
    }
  }
  return null
}
