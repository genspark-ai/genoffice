import { mp4VideoSize } from './mp4-audio-sniff'

export interface Size {
  width: number
  height: number
}

/** Frame size of a video container we know how to embed; null when unreadable. */
export function videoSize(bytes: Uint8Array, ext: string): Size | null {
  switch (ext.toLowerCase()) {
    case 'mp4':
    case 'm4v':
    case 'mov':
      return mp4VideoSize(bytes)
    case 'webm':
    case 'mkv':
      return webmVideoSize(bytes)
    case 'avi':
      return aviVideoSize(bytes)
    default:
      return null
  }
}

const EMU_PER_PX = 12700

/**
 * PowerPoint placement for an inserted video (measured against PowerPoint for Mac):
 * the frame comes in at its pixel size with 1 px = 1 pt, is shrunk to fit inside
 * the slide when it overflows in either direction, never enlarged, and centered.
 * Without a readable size we fall back to a 16:9 frame fitted to the slide.
 * A drop point replaces the slide center as the anchor.
 */
export function videoFrame(
  slide: Size,
  natural: Size | null,
  center?: Point,
): {
  x: number
  y: number
  cx: number
  cy: number
} {
  const known = natural && natural.width > 0 && natural.height > 0
  const w = known ? natural.width * EMU_PER_PX : slide.width
  const h = known ? natural.height * EMU_PER_PX : (slide.width * 9) / 16
  const scale = Math.min(1, slide.width / w, slide.height / h)
  return centeredFrame(slide, Math.round(w * scale), Math.round(h * scale), center)
}

const AUDIO_ICON_EMU = 64 * EMU_PER_PX

/** PowerPoint drops an inserted audio file as a 64 pt speaker icon centered on the slide, whatever the slide size. */
export function audioFrame(
  slide: Size,
  center?: Point,
): { x: number; y: number; cx: number; cy: number } {
  return centeredFrame(slide, AUDIO_ICON_EMU, AUDIO_ICON_EMU, center)
}

export interface Point {
  x: number
  y: number
}

function centeredFrame(slide: Size, cx: number, cy: number, center?: Point) {
  const c = center ?? { x: slide.width / 2, y: slide.height / 2 }
  return { x: Math.round(c.x - cx / 2), y: Math.round(c.y - cy / 2), cx, cy }
}

// ── EBML (webm/mkv) ──────────────────────────────────────────────────────

const EBML_SEGMENT = 0x18538067
const EBML_TRACKS = 0x1654ae6b
const EBML_TRACK_ENTRY = 0xae
const EBML_VIDEO = 0xe0
const EBML_PIXEL_WIDTH = 0xb0
const EBML_PIXEL_HEIGHT = 0xba
const EBML_DISPLAY_WIDTH = 0x54b0
const EBML_DISPLAY_HEIGHT = 0x54ba
const EBML_CLUSTER = 0x1f43b675
const EBML_CONTAINERS = new Set([EBML_SEGMENT, EBML_TRACKS, EBML_TRACK_ENTRY, EBML_VIDEO])

/** Reads a vint; ids keep the length marker, sizes strip it. All-ones size = unknown (-1). */
function vint(bytes: Uint8Array, off: number, end: number, keepMarker: boolean) {
  if (off >= end) return null
  const first = bytes[off]!
  let len = 1
  while (len <= 8 && !(first & (0x80 >> (len - 1)))) len++
  if (len > 8 || off + len > end) return null
  let value = keepMarker ? first : first & (0xff >> len)
  let allOnes = !keepMarker && value === 0xff >> len
  for (let i = 1; i < len; i++) {
    const b = bytes[off + i]!
    value = value * 256 + b
    if (b !== 0xff) allOnes = false
  }
  return { value: allOnes ? -1 : value, len }
}

function readUint(bytes: Uint8Array, off: number, len: number): number {
  let v = 0
  for (let i = 0; i < len; i++) v = v * 256 + bytes[off + i]!
  return v
}

function webmVideoSize(bytes: Uint8Array): Size | null {
  const found: Partial<Record<'pw' | 'ph' | 'dw' | 'dh', number>> = {}
  const scan = (start: number, end: number): boolean => {
    let off = start
    while (off < end) {
      const id = vint(bytes, off, end, true)
      if (!id) return false
      const size = vint(bytes, off + id.len, end, false)
      if (!size) return false
      const payload = off + id.len + size.len
      const boxEnd = size.value < 0 ? end : Math.min(end, payload + size.value)
      if (id.value === EBML_CLUSTER) return false
      if (EBML_CONTAINERS.has(id.value)) {
        if (scan(payload, boxEnd)) return true
        if (id.value === EBML_VIDEO && found.pw && found.ph) return true
      } else if (size.value >= 0 && size.value <= 8) {
        const v = readUint(bytes, payload, size.value)
        if (id.value === EBML_PIXEL_WIDTH) found.pw = v
        else if (id.value === EBML_PIXEL_HEIGHT) found.ph = v
        else if (id.value === EBML_DISPLAY_WIDTH) found.dw = v
        else if (id.value === EBML_DISPLAY_HEIGHT) found.dh = v
      }
      if (size.value < 0) return false
      off = boxEnd
    }
    return false
  }
  scan(0, bytes.length)
  // DisplayWidth/Height carry the intended aspect (anamorphic sources); both must be present
  const width = found.dw && found.dh ? found.dw : found.pw
  const height = found.dw && found.dh ? found.dh : found.ph
  return width && height ? { width, height } : null
}

// ── RIFF (avi) ───────────────────────────────────────────────────────────

function readU32le(bytes: Uint8Array, off: number): number {
  return (
    (bytes[off]! | (bytes[off + 1]! << 8) | (bytes[off + 2]! << 16) | (bytes[off + 3]! << 24)) >>> 0
  )
}

function tag(bytes: Uint8Array, off: number): string {
  return String.fromCharCode(bytes[off]!, bytes[off + 1]!, bytes[off + 2]!, bytes[off + 3]!)
}

function aviVideoSize(bytes: Uint8Array): Size | null {
  if (bytes.length < 12 || tag(bytes, 0) !== 'RIFF' || tag(bytes, 8) !== 'AVI ') return null
  let off = 12
  const end = Math.min(bytes.length, 8 + readU32le(bytes, 4))
  while (off + 8 <= end) {
    const size = readU32le(bytes, off + 4)
    const type = tag(bytes, off)
    if (type === 'LIST' && off + 12 <= end && tag(bytes, off + 8) === 'hdrl') {
      const p = off + 12
      // avih payload: dwWidth at +32, dwHeight at +36
      if (p + 48 <= end && tag(bytes, p) === 'avih') {
        const width = readU32le(bytes, p + 40)
        const height = readU32le(bytes, p + 44)
        return width && height ? { width, height } : null
      }
      return null
    }
    off += 8 + size + (size & 1)
  }
  return null
}
