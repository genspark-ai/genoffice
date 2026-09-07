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

/**
 * Drop abandoned page-turn experiments from a local snapshot. Stock studio
 * behavior is restored; only `file:new-doc` stays patched.
 */
export function stripAbandonedStudioPatches(js) {
  let next = js
  next = next.replace(
    /\/\*genoffice-eager-prefetch\*\/n\(\)/g,
    'if(typeof r.requestIdleCallback==`function`){this.deferredPrefetchTask={kind:`idle`,id:r.requestIdleCallback(n,{timeout:1e3})};return}this.deferredPrefetchTask={kind:`timeout`,id:window.setTimeout(n,250)}',
  )
  next = next.replace(
    /\/\*genoffice-prefetch-overscan\*\/for\(let e of\[s-2,s-1,c\+1,c\+2\]\)/g,
    'for(let e of[s-1,c+1])',
  )
  next = next.replace(
    /\/\*genoffice-page-align\*\/n>0\?([A-Za-z_$][\w$]*)\(e,r,o,a\):([A-Za-z_$][\w$]*)\(e,r,o\)/g,
    'n>0?Math.min($1(e,r,o,a),r+i):Math.max($2(e,r,o),r-i)',
  )
  next = next.replace(
    /flushDeferredPaginationIfNeeded\(`before-navigation`,\/\*genoffice-nav-pagination\*\/!0\)/g,
    'flushDeferredPaginationIfNeeded(`before-navigation`,!1)',
  )
  next = next.replace(
    /e!==`document-agent-rendered`&&\/\*genoffice-sync-layout\*\/\(typeof e==`string`&&e\.includes\(`deferred-pagination-flush`\)\?this\.refreshPages\(\):this\.refreshPagesForMutation\(\)\)/g,
    'e!==`document-agent-rendered`&&this.refreshPagesForMutation()',
  )
  next = next.replace(
    /this\.recalcLayout\(\),\/\*genoffice-clamp-scroll\*\/\(\(\)=>\{let e=this\.viewportManager\.getViewportSize\(\),t=Math\.max\(0,this\.virtualScroll\.getTotalHeight\(\)-e\.height\);this\.viewportManager\.getScrollY\(\)>t&&this\.viewportManager\.setScrollTop\(t\)\}\)\(\),this\.cancelPendingTextEditRefresh\(\)/g,
    'this.recalcLayout(),this.cancelPendingTextEditRefresh()',
  )
  next = next.replace(
    /vp\.call\(this,e\.key===`PageUp`\?-1:1,e\.shiftKey\),\/\*genoffice-caret-after-page\*\/this\.updateCaret\(\);return\}/g,
    'vp.call(this,e.key===`PageUp`?-1:1,e.shiftKey);return}',
  )
  return next
}
