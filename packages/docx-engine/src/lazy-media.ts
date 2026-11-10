/**
 * Lazy media: a large document's browser-renderable pictures stay in the file
 * on disk instead of travelling into the renderer. The zip handed to the
 * renderer carries a fixed-size placeholder per stripped part naming the
 * source document's sha256; the picture is then referenced by URL and served
 * from the original file, and a save swaps the placeholders back for the
 * original bytes (see zip-splice.ts).
 */

export const LAZY_MEDIA_SCHEME = 'genoffice-docx-media'

const MAGIC = 'GENOFFICE-LAZY-MEDIA\n'
const HASH_RE = /^[0-9a-f]{64}$/

export const LAZY_MEDIA_PLACEHOLDER_BYTES = MAGIC.length + 64 + 1

export function lazyMediaPlaceholder(hash: string): Uint8Array {
  if (!HASH_RE.test(hash)) throw new Error(`lazy media: not a sha256 hex: ${hash}`)
  return new TextEncoder().encode(`${MAGIC}${hash}\n`)
}

export function lazyMediaHashOf(bytes: Uint8Array): string | null {
  if (bytes.length !== LAZY_MEDIA_PLACEHOLDER_BYTES) return null
  const text = new TextDecoder('latin1').decode(bytes)
  if (!text.startsWith(MAGIC) || !text.endsWith('\n')) return null
  const hash = text.slice(MAGIC.length, MAGIC.length + 64)
  return HASH_RE.test(hash) ? hash : null
}

/** pictures Chromium decodes itself; emf/wmf/tiff need their bytes in the renderer */
export function isLazyMediaPart(partPath: string): boolean {
  return /^word\/media\/[\x21-\x7e]+\.(?:png|jpe?g|gif|bmp|webp|svg)$/i.test(partPath)
}

/** Max lazy-media part path chars served from the source archive. */
export const MAX_LAZY_PART_CHARS = 512

export function lazyMediaUrl(hash: string, partPath: string): string {
  if (!HASH_RE.test(hash)) throw new Error(`lazy media: not a sha256 hex: ${hash}`)
  return `${LAZY_MEDIA_SCHEME}://${hash}/${partPath.split('/').map(encodeURIComponent).join('/')}`
}

export function parseLazyMediaUrl(url: string): { hash: string; partPath: string } | null {
  const m = new RegExp(`^${LAZY_MEDIA_SCHEME}://([0-9a-f]{64})/(.+)$`).exec(url)
  if (!m) return null
  let partPath: string
  try {
    partPath = m[2].split('/').map(decodeURIComponent).join('/')
  } catch {
    return null
  }
  if (partPath.length > MAX_LAZY_PART_CHARS || partPath.includes('\0')) return null
  if (partPath.split('/').some((seg) => seg === '..')) return null
  return { hash: m[1], partPath }
}
