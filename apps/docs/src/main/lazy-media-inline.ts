import { LAZY_MEDIA_SCHEME } from '@genoffice/docx-engine/lazy-media'

/**
 * Exported HTML leaves the app: every lazily served picture it references
 * must travel as bytes. Distinct URLs are read once; one the source cannot
 * serve any more keeps its reference (the export still opens, that picture
 * shows as missing, as it would in the editor).
 */
export async function inlineLazyMediaInHtml(
  html: string,
  read: (url: string) => Promise<{ body: Uint8Array; mime: string } | null>,
): Promise<string> {
  // part names may carry apostrophes (encodeURIComponent keeps them): the URL
  // runs to the closing double quote, whitespace or a tag bracket
  const re = new RegExp(`${LAZY_MEDIA_SCHEME}://[^"\\s<>]+`, 'g')
  const urls = [...new Set(html.match(re) ?? [])]
  if (urls.length === 0) return html
  const dataUrls = new Map<string, string>()
  for (const url of urls) {
    const media = await read(url)
    if (media) {
      dataUrls.set(url, `data:${media.mime};base64,${Buffer.from(media.body).toString('base64')}`)
    }
  }
  return html.replace(re, (url) => dataUrls.get(url) ?? url)
}
