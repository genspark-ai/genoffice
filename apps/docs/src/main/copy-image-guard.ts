/**
 * Pure guards for docs:copy-image-to-clipboard. The data URL is base64-decoded
 * and rasterized in the main process, so the payload is capped before decode
 * and only raster subtypes are accepted. Electron-free for direct unit tests.
 */

/** Renderer re-encodes via canvas.toDataURL, so photos routinely exceed 15 MB. */
export const MAX_COPY_IMAGE_DATA_URL = 64 * 1024 * 1024

const DATA_URL_RE = /^data:image\/(png|jpe?g|gif|webp|bmp);/i

export function validCopyImageDataUrl(dataUrl: unknown): dataUrl is string {
  if (typeof dataUrl !== 'string') return false
  if (dataUrl.length > MAX_COPY_IMAGE_DATA_URL) return false
  // Non-data URLs are lazy-media ids resolved server-side; data URLs must be
  // raster images (never data:text/html or svg scripts).
  if (dataUrl.startsWith('data:') && !DATA_URL_RE.test(dataUrl)) return false
  return true
}

/** Display-size overrides for the pasted <img>: finite, clamped to 100000px. */
export function copyImageDisplaySize(meta: unknown): { width?: number; height?: number } {
  if (typeof meta !== 'string' || meta.length > 2048) return {}
  try {
    const parsed = JSON.parse(meta) as Record<string, unknown>
    const out: { width?: number; height?: number } = {}
    if (
      typeof parsed.imageWidthPx === 'number' &&
      Number.isFinite(parsed.imageWidthPx) &&
      parsed.imageWidthPx > 0
    ) {
      out.width = Math.min(Math.round(parsed.imageWidthPx), 100000)
    }
    if (
      typeof parsed.imageHeightPx === 'number' &&
      Number.isFinite(parsed.imageHeightPx) &&
      parsed.imageHeightPx > 0
    ) {
      out.height = Math.min(Math.round(parsed.imageHeightPx), 100000)
    }
    return out
  } catch {
    return {}
  }
}
