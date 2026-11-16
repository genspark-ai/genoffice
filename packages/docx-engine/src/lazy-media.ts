/**
 * Lazy media: a large document's browser-renderable pictures stay in the file
 * on disk instead of travelling into the renderer. The zip handed to the
 * renderer carries a fixed-size placeholder per stripped part naming the
 * source document's sha256; the picture is then referenced by URL and served
 * from the original file, and a save swaps the placeholders back for the
 * original bytes (see zip-splice.ts).
 */

import { DEFAULT_ZIP_LIMITS, inflateRawWithinCap } from '@genoffice/zip-gate'

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

/** One media part as the source archive's central directory describes it. */
export interface LazyMediaPartRef {
  method: number
  csize: number
  usize: number
  dataOffset: number
}

/**
 * Memory budget for serving one media part out of the source archive.
 *
 * The lazy path skips the bulk open's metered-inflate gate (`zip-load` runs
 * it, `openZipFile`/`readZipEntries` do not — their pass reads declared
 * sizes only), so serving must hold on its own: both header words are forged
 * at will by a hostile archive, and the inflate would otherwise pay for
 * whatever the payload really expands to (measured in review, 2026-11-14:
 * 299 KB archive in, 300 MB inflated out, arbitrary scale). Aligned with the
 * bulk gate's per-part cap, so any archive the bulk gate accepts also serves.
 */
export const MAX_LAZY_SERVED_MEDIA_BYTES = DEFAULT_ZIP_LIMITS.maxPartBytes

/**
 * The bytes of one media part from the source archive, bounded.
 *
 * `readRange` allocates the compressed range in one buffer from the
 * archive's own word, so the csize/usize headers are checked before it is
 * called; the inflate is then cancelled past the cap, returning null. Null
 * also covers unsupported compression methods and corrupt streams — the
 * serving protocol handler maps every null to a 404.
 */
export async function readMediaPartBounded(
  part: LazyMediaPartRef,
  readRange: (offset: number, length: number) => Promise<Uint8Array>,
  maxBytes: number = MAX_LAZY_SERVED_MEDIA_BYTES,
): Promise<Uint8Array | null> {
  if (part.method !== 0 && part.method !== 8) return null
  if (part.csize > maxBytes || part.usize > maxBytes) return null
  const raw = await readRange(part.dataOffset, part.csize)
  if (part.method === 0) return raw
  return inflateRawWithinCap(raw, maxBytes)
}
