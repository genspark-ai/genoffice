/**
 * Neutralize a JPEG's EXIF Orientation in place (tag 0x0112 -> 1).
 *
 * PowerPoint ignores EXIF orientation and renders the raw pixel grid; Chromium
 * auto-applies it when decoding (`image-orientation: from-image` has no canvas
 * opt-out via data URLs). Files that combine a rotated pixel grid with an EXIF
 * flag AND a shape-level rot therefore render 90 degrees off. Rewriting the flag
 * byte before serving keeps every consumer (canvas, thumbnails, exports) on
 * PowerPoint's semantics; the archive keeps the original bytes for save fidelity.
 */
export function neutralizeJpegOrientation(bytes: Uint8Array): Uint8Array {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return bytes
  let off = 2
  while (off + 4 <= bytes.length) {
    if (bytes[off] !== 0xff) return bytes
    const marker = bytes[off + 1]!
    // Start of scan / end: no APP1 past this point
    if (marker === 0xda || marker === 0xd9) return bytes
    const segLen = (bytes[off + 2]! << 8) | bytes[off + 3]!
    if (segLen < 2 || off + 2 + segLen > bytes.length) return bytes
    if (marker === 0xe1) {
      const p = off + 4
      // "Exif\0\0"
      if (
        bytes[p] === 0x45 &&
        bytes[p + 1] === 0x78 &&
        bytes[p + 2] === 0x69 &&
        bytes[p + 3] === 0x66 &&
        bytes[p + 4] === 0x00 &&
        bytes[p + 5] === 0x00
      ) {
        const tiff = p + 6
        const le = bytes[tiff] === 0x49 && bytes[tiff + 1] === 0x49
        const be = bytes[tiff] === 0x4d && bytes[tiff + 1] === 0x4d
        if (!le && !be) return bytes
        const u16 = (o: number) =>
          le ? bytes[o]! | (bytes[o + 1]! << 8) : (bytes[o]! << 8) | bytes[o + 1]!
        const u32 = (o: number) =>
          le
            ? (bytes[o]! | (bytes[o + 1]! << 8) | (bytes[o + 2]! << 16) | (bytes[o + 3]! << 24)) >>>
              0
            : ((bytes[o]! << 24) | (bytes[o + 1]! << 16) | (bytes[o + 2]! << 8) | bytes[o + 3]!) >>>
              0
        const segEnd = off + 2 + segLen
        const ifd0 = tiff + u32(tiff + 4)
        if (ifd0 + 2 > segEnd) return bytes
        const count = u16(ifd0)
        for (let i = 0; i < count; i++) {
          const e = ifd0 + 2 + i * 12
          if (e + 12 > segEnd) return bytes
          if (u16(e) !== 0x0112) continue
          const cur = u16(e + 8)
          if (cur === 1 || cur === 0) return bytes
          const out = new Uint8Array(bytes)
          if (le) {
            out[e + 8] = 1
            out[e + 9] = 0
          } else {
            out[e + 8] = 0
            out[e + 9] = 1
          }
          return out
        }
        return bytes
      }
    }
    off += 2 + segLen
  }
  return bytes
}
