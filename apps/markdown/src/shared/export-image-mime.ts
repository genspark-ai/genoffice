/** Every format the editor displays; the DOCX export rasterizes the non-native ones */
export type ExportImageMime =
  | 'image/png'
  | 'image/jpeg'
  | 'image/gif'
  | 'image/webp'
  | 'image/svg+xml'
  | 'image/bmp'
  | 'image/avif'

export const EXPORT_IMAGE_MIME_BY_EXT: Record<string, ExportImageMime> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
  '.bmp': 'image/bmp',
  '.avif': 'image/avif',
}

export const EXPORT_IMAGE_EXTS = Object.keys(EXPORT_IMAGE_MIME_BY_EXT).map((ext) => ext.slice(1))

const MIME_ALIASES: Record<string, ExportImageMime> = {
  'image/jpg': 'image/jpeg',
  'image/pjpeg': 'image/jpeg',
  'image/svg': 'image/svg+xml',
  'image/x-ms-bmp': 'image/bmp',
  'image/x-bmp': 'image/bmp',
}

export function exportImageMime(mime: string | null | undefined): ExportImageMime | null {
  const bare = String(mime ?? '')
    .split(';')[0]!
    .trim()
    .toLowerCase()
  if (!bare) return null
  if (MIME_ALIASES[bare]) return MIME_ALIASES[bare]
  return (Object.values(EXPORT_IMAGE_MIME_BY_EXT) as string[]).includes(bare)
    ? (bare as ExportImageMime)
    : null
}

export function exportImageMimeFromPath(path: string): ExportImageMime | null {
  const clean = path.split(/[?#]/)[0]!
  const dot = clean.lastIndexOf('.')
  if (dot < 0) return null
  return EXPORT_IMAGE_MIME_BY_EXT[clean.slice(dot).toLowerCase()] ?? null
}

function startsWith(bytes: Uint8Array, sig: number[], offset = 0): boolean {
  if (bytes.length < offset + sig.length) return false
  for (let i = 0; i < sig.length; i++) if (bytes[offset + i] !== sig[i]) return false
  return true
}

/** Container signature of the bytes; SVG has none, so it falls to the header/extension */
export function sniffExportImageMime(bytes: Uint8Array): ExportImageMime | null {
  if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47])) return 'image/png'
  if (startsWith(bytes, [0xff, 0xd8, 0xff])) return 'image/jpeg'
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) return 'image/gif'
  if (startsWith(bytes, [0x42, 0x4d])) return 'image/bmp'
  if (startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) && startsWith(bytes, [0x57, 0x45, 0x42, 0x50], 8))
    return 'image/webp'
  if (startsWith(bytes, [0x66, 0x74, 0x79, 0x70, 0x61, 0x76, 0x69], 4)) return 'image/avif'
  return null
}

/**
 * Format of a downloaded picture: the bytes decide when they carry a
 * signature, else the Content-Type header, else the URL's extension.
 */
export function remoteExportImageMime(
  bytes: Uint8Array,
  contentType: string | null,
  url: string,
): ExportImageMime | null {
  return sniffExportImageMime(bytes) ?? exportImageMime(contentType) ?? exportImageMimeFromPath(url)
}
