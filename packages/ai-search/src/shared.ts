/** Search result types and shared constants (used by both the index and gsk backends) */

export interface WebSearchResult {
  title: string
  url: string
  snippet: string
}

export interface ImageSearchResult {
  title: string
  imageUrl: string
  sourceUrl: string
  source: string
  width?: number
  height?: number
}

// Known stock-photo hosts skipped during image search (matches the upstream filter list)
export const COPYRIGHT_HOSTS = ['gettyimages', 'istockphoto', 'shutterstock', 'corbis']

/**
 * True when an image URL lives on a stock-photo host. Scoped to the hostname
 * (not a full-URL substring): a blog image whose *path* merely mentions a
 * stock site ("…/shutterstock-review.png") is kept, while host matching keeps
 * the previous behavior (gettyimages.com and its subdomains stay blocked).
 */
export function isCopyrightHost(imageUrl: string): boolean {
  const host = safeHost(imageUrl).toLowerCase()
  if (!host) return false
  return COPYRIGHT_HOSTS.some((d) => host.includes(d))
}

export function safeHost(url: unknown): string {
  try {
    return new URL(String(url)).hostname
  } catch {
    return ''
  }
}

/**
 * View untrusted JSON as a string-keyed record so properties can be probed
 * without `any`; non-object inputs read as an empty record.
 */
export function asRecord(v: unknown): Record<string, unknown> {
  return typeof v === 'object' && v !== null ? (v as Record<string, unknown>) : {}
}

/** First element when the value is an array, otherwise undefined (loose JSON probing). */
export function firstItem(v: unknown): unknown {
  return Array.isArray(v) ? (v as unknown[])[0] : undefined
}

let explicitProxyUrl = ''

/**
 * Proxy resolved by the apps' proxy bootstraps (env vars, else the system
 * proxy via session.resolveProxy); consumed by gskChildEnv() and the login
 * flow's proxy fallback.
 */
export function setGskProxyUrl(url: string): void {
  explicitProxyUrl = url
}

export function gskProxyUrl(): string {
  return explicitProxyUrl
}
