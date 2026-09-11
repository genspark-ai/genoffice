/** Same-origin sibling of the vendored studio (`/rhwp/print.html`). */
const PRINT_SURFACE_PATH = '/rhwp/print.html'

export function isHwpPrintSurfaceUrl(openerUrl: string, targetUrl: string): boolean {
  try {
    const opener = new URL(openerUrl)
    const target = new URL(targetUrl, opener)
    if (target.origin !== opener.origin) return false
    if (target.protocol !== 'http:' && target.protocol !== 'https:') return false
    return target.pathname.replace(/\/+$/, '') === PRINT_SURFACE_PATH
  } catch {
    return false
  }
}

/** Absolute print.html URL, or null when the popup is already on that page. */
export function printSurfaceLoadUrl(
  openerUrl: string,
  requestedUrl: string,
  currentUrl: string,
): string | null {
  if (!isHwpPrintSurfaceUrl(openerUrl, requestedUrl)) return null
  if (isHwpPrintSurfaceUrl(openerUrl, currentUrl)) return null
  return new URL(requestedUrl, openerUrl).href
}
