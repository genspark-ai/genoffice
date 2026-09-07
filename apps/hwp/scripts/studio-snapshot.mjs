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
