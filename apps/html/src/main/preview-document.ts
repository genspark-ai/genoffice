export const PREVIEW_SCHEME = 'html-preview'
export const ASSET_SCHEME = 'html-asset'

/** Encode an absolute directory as an html-asset base URL with a trailing slash.
 * The scheme is registered as standard, so it needs a host; `local` is a fixed placeholder. */
export function assetBaseHref(documentDir: string): string {
  const normalized = documentDir.replace(/\\/g, '/').replace(/\/+$/, '')
  const path = normalized.startsWith('/') ? normalized : `/${normalized}`
  const encoded = path
    .split('/')
    .map((segment) => encodeURIComponent(segment).replace(/%3A/gi, ':'))
    .join('/')
  return `${ASSET_SCHEME}://local${encoded}/`
}

/**
 * The preview copy of the document: the buffer text with a <base> pointing at
 * the document's directory (so relative assets resolve through html-asset://)
 * unless the author already declared one. The saved file never contains it.
 */
export function buildPreviewDocument(text: string, baseHref: string | null): string {
  if (!baseHref || /<base\s/i.test(text)) return text
  const tag = `<base href="${baseHref.replace(/"/g, '%22')}">`
  const head = /<head(?:\s[^>]*)?>/i.exec(text)
  if (head) {
    const at = head.index + head[0].length
    return text.slice(0, at) + tag + text.slice(at)
  }
  const html = /<html(?:\s[^>]*)?>/i.exec(text)
  if (html) {
    const at = html.index + html[0].length
    return text.slice(0, at) + tag + text.slice(at)
  }
  return tag + text
}

export function previewUrlFor(webContentsId: number): string {
  return `${PREVIEW_SCHEME}://view-${webContentsId}/`
}
