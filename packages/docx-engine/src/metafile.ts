import { convertEmfToDataUrl, convertWmfToDataUrl } from './vendor/emf-converter/index.mjs'

const EMF_MIMES = new Set(['image/emf', 'image/x-emf'])
const WMF_MIMES = new Set(['image/wmf', 'image/x-wmf'])

export function isMetafileMime(mime: string | undefined): mime is string {
  return mime !== undefined && (EMF_MIMES.has(mime) || WMF_MIMES.has(mime))
}

/**
 * Render EMF/WMF bytes to a PNG data URL via the vendored emf-converter.
 * Returns null on parse failure or when no canvas API exists (non-renderer
 * environments), so callers keep their existing empty-frame degrade.
 */
export async function metafileToDataUrl(
  bytes: ArrayBuffer | Uint8Array,
  mime: string,
): Promise<string | null> {
  const buffer = bytes instanceof Uint8Array ? bytes.slice().buffer : bytes
  try {
    if (EMF_MIMES.has(mime)) return await convertEmfToDataUrl(buffer, { dpiScale: 2 })
    if (WMF_MIMES.has(mime)) return await convertWmfToDataUrl(buffer, { dpiScale: 2 })
    return null
  } catch {
    return null
  }
}
