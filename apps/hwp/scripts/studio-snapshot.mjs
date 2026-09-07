/** Shared rhwp-studio snapshot helpers for the vendor script and tests. */

export const PWA_FILES = ['sw.js', 'registerSW.js', 'manifest.webmanifest']

export const REQUIRED_RELATIVE = [
  'index.html',
  'fonts/NotoSansKR-Regular.woff2',
  'fonts/Pretendard-Regular.woff2',
]

export const REQUIRED_ASSET_EXTS = ['.js', '.wasm']

const PWA_HTML_RE =
  /<link\s+rel="manifest"[^>]*>|<script[^>]*(?:id="vite-plugin-pwa:register-sw"|src="[^"]*registerSW\.js")[^>]*><\/script>/g

export function isPwaPath(urlPath) {
  const name = urlPath.split('?')[0].split('/').pop() ?? ''
  return PWA_FILES.includes(name) || /^workbox-.*\.js$/.test(name)
}

export function stripPwaHtml(html) {
  return html.replace(PWA_HTML_RE, '')
}

/** Stock studio defers neighbor-page raster until idle (up to 1s). Marker written after patch. */
export const EAGER_PREFETCH_MARK = '/*genoffice-eager-prefetch*/'

const IDLE_PREFETCH_RE =
  /if\(typeof r\.requestIdleCallback==`function`\)\{this\.deferredPrefetchTask=\{kind:`idle`,id:r\.requestIdleCallback\(n,\{timeout:1e3\}\)\};return\}this\.deferredPrefetchTask=\{kind:`timeout`,id:window\.setTimeout\(n,250\)\}/

/**
 * Paint the next page as soon as the current one is on screen.
 * Neighbor first-paint is cheap enough; 0.8.x still idles the same prefetch, so we run it
 * immediately instead of waiting for requestIdleCallback.
 */
export function eagerPagePrefetch(js) {
  if (js.includes(EAGER_PREFETCH_MARK)) return js
  if (!js.includes('schedulePrefetchPages') || !js.includes('requestIdleCallback')) return js
  const next = js.replace(IDLE_PREFETCH_RE, `${EAGER_PREFETCH_MARK}n()`)
  if (next === js) {
    throw new Error('rhwp-studio prefetch idle deferral changed — update eagerPagePrefetch()')
  }
  return next
}

export function hasEagerPrefetch(js) {
  return js.includes(EAGER_PREFETCH_MARK)
}

/**
 * Embed mode strips File new/open/save from the registry so the host owns those
 * actions — and also skips boot-time `createNewDocument()`. Untitled tabs then
 * have no pages. Keep only `file:new-doc` registered; OA() still hides the menu.
 */
export const EMBED_NEW_DOC_MARK = '/*genoffice-embed-new-doc*/'

const EMBED_NEW_DOC_RE =
  /bA\.registerAll\(yA===`embed`\?Ev\.filter\(e=>!sD\.includes\(e\.id\)\):Ev\)/

export function keepEmbedNewDoc(js) {
  if (js.includes(EMBED_NEW_DOC_MARK)) return js
  if (!js.includes('file:new-doc') || !js.includes('registerAll')) return js
  const next = js.replace(
    EMBED_NEW_DOC_RE,
    `bA.registerAll(yA===\`embed\`?Ev.filter(e=>${EMBED_NEW_DOC_MARK}e.id===\`file:new-doc\`||!sD.includes(e.id)):Ev)`,
  )
  if (next === js) {
    throw new Error('rhwp-studio embed command filter changed — update keepEmbedNewDoc()')
  }
  return next
}

export function hasEmbedNewDoc(js) {
  return js.includes(EMBED_NEW_DOC_MARK)
}
