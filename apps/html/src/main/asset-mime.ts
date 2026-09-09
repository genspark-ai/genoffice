import { basename, extname } from 'node:path'
import type { ImageData } from '../shared/ipc'

/** Sibling files the preview may load through html-asset:// by extension.
 * Deliberately no json / txt / html: a page's own script must not be able to
 * read arbitrary neighbours and exfiltrate them. */
export const PREVIEW_ASSET_EXTS = new Set([
  '.png',
  '.jpg',
  '.jpeg',
  '.gif',
  '.webp',
  '.svg',
  '.bmp',
  '.avif',
  '.ico',
  '.css',
  '.js',
  '.mjs',
  '.woff',
  '.woff2',
  '.ttf',
  '.otf',
  '.mp4',
  '.webm',
  '.mp3',
  '.ogg',
  '.wav',
])

export const ASSET_SNIFF_BYTES = 64

function ascii(head: Uint8Array, start: number, length: number): string {
  return String.fromCharCode(...head.subarray(start, start + length))
}

function startsWith(head: Uint8Array, bytes: number[], offset = 0): boolean {
  if (head.length < offset + bytes.length) return false
  return bytes.every((b, i) => head[offset + i] === b)
}

/** Binary signatures only: a file that starts like a JPEG is not a secret. */
export function sniffBinaryAssetMime(head: Uint8Array): string | null {
  if (startsWith(head, [0xff, 0xd8, 0xff])) return 'image/jpeg'
  if (startsWith(head, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return 'image/png'
  if (head.length >= 6 && /^GIF8[79]a$/.test(ascii(head, 0, 6))) return 'image/gif'
  if (head.length >= 12 && ascii(head, 0, 4) === 'RIFF') {
    const form = ascii(head, 8, 4)
    if (form === 'WEBP') return 'image/webp'
    if (form === 'WAVE') return 'audio/wav'
    return null
  }
  if (startsWith(head, [0x42, 0x4d]) && startsWith(head, [0, 0, 0, 0], 6)) return 'image/bmp'
  if (startsWith(head, [0x00, 0x00, 0x01, 0x00])) return 'image/x-icon'
  if (head.length >= 12 && ascii(head, 4, 4) === 'ftyp') {
    return /^avi[fs]$/.test(ascii(head, 8, 4)) ? 'image/avif' : 'video/mp4'
  }
  if (startsWith(head, [0x1a, 0x45, 0xdf, 0xa3])) return 'video/webm'
  if (head.length >= 4 && ascii(head, 0, 4) === 'OggS') return 'audio/ogg'
  if (head.length >= 3 && ascii(head, 0, 3) === 'ID3') return 'audio/mpeg'
  if (head.length >= 2 && head[0] === 0xff && (head[1]! & 0xe6) === 0xe2) return 'audio/mpeg'
  if (head.length >= 4) {
    const tag = ascii(head, 0, 4)
    if (tag === 'wOFF') return 'font/woff'
    if (tag === 'wOF2') return 'font/woff2'
    if (tag === 'OTTO') return 'font/otf'
    if (tag === 'true' || startsWith(head, [0x00, 0x01, 0x00, 0x00])) {
      const numTables = head.length >= 6 ? (head[4]! << 8) | head[5]! : 0
      if (numTables > 0 && numTables <= 64) return 'font/ttf'
    }
  }
  return null
}

function looksLikeSvg(head: Uint8Array): boolean {
  const text = new TextDecoder('utf-8', { fatal: false }).decode(head).replace(/^\uFEFF/, '')
  return /^\s*<(\?xml|svg[\s>])/i.test(text)
}

export type AssetSlot = 'image' | 'style' | 'other'

/** Electron strips Sec-Fetch-* before protocol.handle sees a request, so the
 * slot is read from Accept: images ask for `image/...`, stylesheets for
 * `text/css`, everything else (script, font, media, fetch) sends `*\/*`. */
export function slotFromAccept(accept: string | null): AssetSlot {
  const first = (accept ?? '').trim().toLowerCase()
  if (first.startsWith('image/')) return 'image'
  if (first.startsWith('text/css')) return 'style'
  return 'other'
}

const TEXT_MIME_BY_SLOT: Partial<Record<AssetSlot, string>> = {
  style: 'text/css; charset=utf-8',
  other: 'text/javascript; charset=utf-8',
}

/**
 * Content type for a sibling file whose extension is not on the allowlist.
 * "Save page as, complete" in Chrome/Safari writes images, stylesheets and
 * scripts without extensions (`_files/tTsEQygh`), so binary media is sniffed
 * from its signature and text assets are typed by the slot that asked for
 * them. A sniffed image is only served to an image slot and other binaries
 * only to a `*\/*` slot; text is only served when the file has no extension
 * at all and is not a dotfile, so `.json` / `.txt` / `.env` neighbours stay
 * unreadable. Responses stay cross-origin and opaque to page script either
 * way. Returns null when the file must not be served.
 */
export function extensionlessAssetMime(
  target: string,
  head: Uint8Array,
  accept: string | null,
): string | null {
  const slot = slotFromAccept(accept)
  const binary = sniffBinaryAssetMime(head)
  if (binary)
    return (binary.startsWith('image/') ? slot === 'image' : slot === 'other') ? binary : null
  const name = basename(target)
  if (extname(name) !== '' || name.startsWith('.')) return null
  if (slot === 'image') return looksLikeSvg(head) ? 'image/svg+xml' : null
  return TEXT_MIME_BY_SLOT[slot] ?? null
}

/** data-URL label for a picture read back for editing: PNG / GIF keep their type, any other image
 * signature is labelled JPEG the way fetchImage does (the decoder sniffs the bytes anyway) */
export function editableImageMime(head: Uint8Array): ImageData['mime'] | null {
  const mime = sniffBinaryAssetMime(head)
  if (!mime?.startsWith('image/')) return null
  return mime === 'image/png' || mime === 'image/gif' ? mime : 'image/jpeg'
}
