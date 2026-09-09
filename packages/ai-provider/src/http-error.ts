/**
 * Detail text for a non-OK HTTP response body. Real API errors (JSON/plain
 * text) are surfaced as-is, truncated. HTML bodies — an edge/WAF block page or
 * a login page served instead of an API response (e.g. Genspark's SPA shell on
 * a 403) — are replaced with a short readable note, since dumping raw markup
 * into the chat UI is useless to the user.
 */
export function httpBodyDetail(body: string): string {
  const head = body.trimStart().slice(0, 30).toLowerCase()
  const isHtml = ['<!doctype', '<html', '<head', '<body'].some((tag) => head.startsWith(tag))
  if (isHtml) {
    return `the service returned a web page${describeBlockPage(body)} instead of an API response (likely a temporary network or gateway block) — check your connection and retry`
  }
  return body.slice(0, 500)
}

/** page title plus Cloudflare's error code / Ray ID when present: enough to tell who blocked the call */
function describeBlockPage(body: string): string {
  const title = /<title[^>]*>([^<]*)<\/title>/i.exec(body)?.[1]?.replace(/\s+/g, ' ').trim()
  const code = /error code:?\s*(?:<[^>]+>\s*)*(\d{3,5})/i.exec(body)?.[1]
  const ray = /ray id:?\s*(?:<[^>]+>\s*)*([a-f0-9]{16})/i.exec(body)?.[1]
  const parts = [
    title && `"${title.slice(0, 60)}"`,
    code && `error code ${code}`,
    ray && `ray ${ray}`,
  ].filter(Boolean)
  return parts.length ? ` (${parts.join(', ')})` : ''
}
