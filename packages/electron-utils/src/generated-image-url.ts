/**
 * Names of images written by `storeGeneratedImage`: a UUID plus an image
 * extension. Kept separate from the store so a renderer can recognize its own
 * file:// URLs without importing any node builtins.
 */
export const GENERATED_IMAGE_FILE_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.(png|jpg|gif|webp)$/

export const GENERATED_IMAGE_URL_PREFIX = '/generated-images/'

/**
 * A store file:// URL cannot be loaded from a `genoffice-app://` page, so the
 * preview uses a same-origin path the renderer protocol serves instead. The
 * insert pipelines keep the original file:// URL — only the displayed <img>
 * source changes. Anything that is not one of our own store URLs is returned
 * unchanged so ordinary http(s) and data: pictures still render.
 */
export function generatedImageDisplayUrl(url: string): string {
  let parsed: URL
  try {
    parsed = new URL(url)
  } catch {
    return url
  }
  if (parsed.protocol !== 'file:') return url
  let pathname: string
  try {
    pathname = decodeURIComponent(parsed.pathname)
  } catch {
    return url
  }
  const marker = '/genoffice-ai-images/'
  const at = pathname.lastIndexOf(marker)
  if (at < 0) return url
  const name = pathname.slice(at + marker.length)
  return GENERATED_IMAGE_FILE_RE.test(name) ? `${GENERATED_IMAGE_URL_PREFIX}${name}` : url
}
