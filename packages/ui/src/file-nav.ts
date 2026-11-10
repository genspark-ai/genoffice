import type { MarkdownNav } from './Markdown'

/**
 * filenav links: the shared way for AI answers to cite source files as
 * [name](filenav:///abs/path). Any app's panel can render them by spreading
 * createFileNav() into its Markdown navs; clicking one opens the file in the
 * app that owns its extension (a new tab under the shell), whatever the
 * current app is.
 */

export const FILE_NAV_SCHEME = 'filenav://'

/**
 * filenav href for an absolute path, readable as filenav:///abs/path.
 * encodeURIComponent encodes everything the inline-link href grammar rejects
 * (spaces, %, CJK); it leaves '(' and ')' raw, so those are encoded on top —
 * the parser only accepts one level of balanced parens in an href. Slashes
 * are restored last; safe because '%' from the path is itself encoded.
 */
export function fileNavHref(path: string): string {
  const encoded = encodeURIComponent(path)
    .replace(/%2F/g, '/')
    .replace(/\(/g, '%28')
    .replace(/\)/g, '%29')
  return FILE_NAV_SCHEME + encoded
}

/** filenav:///abs/path → /abs/path (decoded); null when the href carries no path */
export function parseFileNavHref(href: string): string | null {
  if (!href.startsWith(FILE_NAV_SCHEME)) return null
  const encoded = href.slice(FILE_NAV_SCHEME.length)
  if (!encoded) return null
  try {
    return decodeURIComponent(encoded)
  } catch {
    // hand-written hrefs may carry raw (unencoded) text — pass it through
    return encoded
  }
}

/** nav entry that opens cited files through the host app's open-channel */
export function createFileNav(openPath: (path: string) => void): MarkdownNav {
  return {
    scheme: FILE_NAV_SCHEME,
    onNavigate: (href) => {
      const path = parseFileNavHref(href)
      if (path !== null) openPath(path)
    },
  }
}
