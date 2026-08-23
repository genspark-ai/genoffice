/**
 * Display mime for an archive media part: magic bytes win over the file extension.
 * Legacy converters routinely mislabel media (a PNG preview stored as .emf routed
 * the bytes into the EMF parser and rendered nothing), so sniff the real container
 * first and only fall back to the extension when the header is unrecognized.
 */

const EXT_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  // Metafiles keep their real mime so the renderer's image loader rasterizes them
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
}

function startsWith(bytes: Uint8Array, sig: number[], offset = 0): boolean {
  if (bytes.length < offset + sig.length) return false
  for (let i = 0; i < sig.length; i++) if (bytes[offset + i] !== sig[i]) return false
  return true
}

/** Sniffed mime from magic bytes, or null when the header is unrecognized. */
export function sniffImageMime(bytes: Uint8Array): string | null {
  if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47])) return 'image/png'
  if (startsWith(bytes, [0xff, 0xd8, 0xff])) return 'image/jpeg'
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) return 'image/gif'
  if (startsWith(bytes, [0x42, 0x4d])) return 'image/bmp'
  if (startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) && startsWith(bytes, [0x57, 0x45, 0x42, 0x50], 8))
    return 'image/webp'
  if (startsWith(bytes, [0x49, 0x49, 0x2a, 0x00]) || startsWith(bytes, [0x4d, 0x4d, 0x00, 0x2a]))
    return 'image/tiff'
  // EMF: record type 1 (EMR_HEADER) + " EMF" signature at offset 40
  if (
    startsWith(bytes, [0x01, 0x00, 0x00, 0x00]) &&
    startsWith(bytes, [0x20, 0x45, 0x4d, 0x46], 40)
  )
    return 'image/x-emf'
  // Placeable WMF
  if (startsWith(bytes, [0xd7, 0xcd, 0xc6, 0x9a])) return 'image/x-wmf'
  return null
}

/** Effective display mime: sniffed container first, then the extension, then PNG. */
export function displayMime(mediaRef: string, bytes: Uint8Array): string {
  const sniffed = sniffImageMime(bytes)
  if (sniffed) return sniffed
  const ext = mediaRef.split('.').pop()?.toLowerCase() ?? ''
  return EXT_MIME[ext] ?? 'image/png'
}
