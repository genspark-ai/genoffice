import { app } from 'electron'
import { existsSync, readFileSync, statSync } from 'node:fs'
import { join, extname } from 'node:path'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { resolveRendererFile } from '@genoffice/electron-utils'
import { readLazyMedia } from '../../../docs/src/main/lazy-media'
import { takeHandoff } from '../../../docs/src/main/byte-handoff'
import { serveMarkdownAsset } from '../../../markdown/src/main/markdown-main'
import { serveHtmlAsset, previewEntryFor } from '../../../html/src/main/html-main'
import { buildPreviewDocument } from '../../../html/src/main/preview-document'
import { SCHEME_REWRITES } from './schemes'

/**
 * Static front end for the web mode: the six editor renderers are the exact
 * bundles electron-vite builds for the packaged app, served over HTTP. Two
 * serve-time adjustments keep them browser-native without touching their
 * sources:
 *
 *  - the per-app custom media schemes (genoffice-docx-media, md-asset,
 *    html-preview, html-asset) are not fetchable from a browser page; the
 *    scheme names are rewritten to same-origin paths (__media, __mdasset,
 *    __preview, __htmlasset) wherever the literal appears, and HTTP routes map
 *    them back onto the very same main-process handlers.
 *  - index.html gets the web bridge client and the app's preload bundle
 *    injected before the renderer module.
 */

export type RendererRoots = Record<string, string>

export function webRendererRoots(): RendererRoots {
  const names = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html']
  const roots: RendererRoots = {}
  for (const name of names) {
    roots[name] = app.isPackaged
      ? join(process.resourcesPath, 'modules', name, 'renderer')
      : join(app.getAppPath(), '..', name, 'out', 'renderer')
  }
  return roots
}

/** literal scheme name → same-origin path prefix; order matters (longest first) */
const MIME: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.wasm': 'application/wasm',
  '.map': 'application/json',
}

const REWRITTEN_APPS = new Set(Object.keys(SCHEME_REWRITES))

/** extensions served as text (and therefore scheme-rewritten); everything
 *  else is bytes — fonts corrupted badly through a utf8 round trip (OTS
 *  misaligned table) */
const TEXT_EXTS = new Set(['.html', '.js', '.mjs', '.css', '.json', '.map', '.svg'])

/** build in a root directory the client bundles live in (see tools/build-web-preloads.mjs) */
export function webBundleDir(): string {
  return app.isPackaged
    ? join(process.resourcesPath, 'web-preloads')
    : join(__dirname, '..', '..', 'web-preloads')
}

function rewriteText(app: string, text: string): string {
  const rules = SCHEME_REWRITES[app]
  if (!rules) return text
  let out = text
  for (const [scheme, prefix] of rules) out = out.split(scheme).join(prefix)
  return out
}

const injectionFor = (app: string): string =>
  `<script src="/__web/client.js"></script><script src="/__web/preload-${app}.js"></script>`

function serveFile(res: ServerResponse, file: string, app: string): void {
  const ext = extname(file).toLowerCase()
  const mime = MIME[ext] ?? 'application/octet-stream'
  const asText = TEXT_EXTS.has(ext)
  let body: string | Buffer = asText ? readFileSync(file, 'utf8') : readFileSync(file)
  const rewritten = asText && REWRITTEN_APPS.has(app)
  if (rewritten) body = rewriteText(app, body as string)
  if (ext === '.html') {
    body = injectIntoHtml(body as string, injectionFor(app))
  }
  res.writeHead(200, {
    'Content-Type': mime,
    // rewritten files must revalidate; binaries are content-hashed by Vite
    'Cache-Control': rewritten ? 'no-store' : 'public, max-age=3600',
  })
  res.end(body)
}

export function injectIntoHtml(html: string, injection: string): string {
  // the served page rewrites this app's media schemes to same-origin paths;
  // their CSP entries (img-src md-asset: …) become invalid sources, so drop
  // them, and widen connect-src with a bare ws: scheme-source — the desktop
  // CSP's ws://localhost:* misses a 127.0.0.1 or LAN origin
  html = html.replace(/(content="[^"]*)\/__(?:media|mdasset|preview|htmlasset):/g, '$1')
  html = html.replace(/connect-src ([^"]*?)(?=")/, (m, list: string) =>
    m.includes(' ws:') ? m : `connect-src ${list} ws:`,
  )
  const moduleScript = /<script[^>]*type="module"[^>]*>/
  const match = moduleScript.exec(html)
  if (match) return html.slice(0, match.index) + injection + html.slice(match.index)
  return html.replace('</body>', `${injection}</body>`)
}

// ---- web-mode media routes (mapped back onto the original handlers) ----

function sendResponse(res: ServerResponse, response: Response): Promise<void> {
  const headers: Record<string, string> = {}
  response.headers.forEach((value, key) => {
    headers[key] = value
  })
  res.writeHead(response.status, headers)
  if (response.body) {
    return response.arrayBuffer().then((buf) => {
      res.end(Buffer.from(buf))
    })
  }
  res.end()
  return Promise.resolve()
}

/** GET /__media://<hash>/<partPath> → docs embedded media (lazy zip reads).
 *  <handoff>/<uuid> shapes are one-shot document-byte transfers the renderer
 *  received inside an IPC reply — strings through the bridge skip the static
 *  rewrite, so the route maps them back onto the same handoff table. */
async function serveDocsMedia(res: ServerResponse, pathname: string): Promise<boolean> {
  const rest = pathname.replace(/^\/__media:\/\//, '')
  if (rest === pathname) return false
  const schemeUrl = `genoffice-docx-media://${rest}`
  const handed = takeHandoff(schemeUrl)
  if (handed) {
    res.writeHead(200, {
      'Content-Type': 'application/octet-stream',
      'Content-Length': String(handed.byteLength),
      'Access-Control-Allow-Origin': '*',
    })
    res.end(handed)
    return true
  }
  const media = await readLazyMedia(schemeUrl)
  if (!media) {
    res.writeHead(404)
    res.end()
    return true
  }
  res.writeHead(200, { 'Content-Type': media.mime, 'Cache-Control': 'no-store' })
  res.end(Buffer.from(media.body))
  return true
}

/** GET /__mdasset://<absolute path> → images next to an open markdown note */
async function serveMarkdownMedia(res: ServerResponse, pathname: string): Promise<boolean> {
  const rest = pathname.replace(/^\/__mdasset:\/\//, '')
  if (rest === pathname) return false
  await sendResponse(res, await serveMarkdownAsset(`md-asset://${rest}`))
  return true
}

/** GET /__htmlasset://local/<encoded path> → assets next to an open html file.
 *  The scheme's fixed `local` host survives the rewrite, so it stays in rest. */
async function serveHtmlMedia(res: ServerResponse, pathname: string): Promise<boolean> {
  const rest = pathname.replace(/^\/__htmlasset:\/\//, '')
  if (rest === pathname) return false
  await sendResponse(res, await serveHtmlAsset(`html-asset://${rest}`))
  return true
}

/** GET /__preview://view-<wcId>/… → the html app's sandboxed preview document */
function servePreview(res: ServerResponse, pathname: string): boolean {
  const match = /^\/__preview:\/\/view-(\d+)\//.exec(pathname)
  if (!match) return false
  const entry = previewEntryFor(Number(match[1]))
  const doc = entry
    ? buildPreviewDocument(entry.text, entry.baseHref)
    : '<!doctype html><title>preview gone</title>'
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' })
  res.end(doc)
  return true
}

// ---- auth ----

const TOKEN_COOKIE = 'genoffice_web_token'

export function tokenValid(req: IncomingMessage, token: string): boolean {
  const url = new URL(req.url ?? '/', 'http://localhost')
  if (url.searchParams.get('token') === token) return true
  const cookies = req.headers.cookie ?? ''
  return cookies.split(/;\s*/).some((c) => c === `${TOKEN_COOKIE}=${token}`)
}

export function tokenCookieHeader(req: IncomingMessage, token: string): string | null {
  const url = new URL(req.url ?? '/', 'http://localhost')
  if (url.searchParams.get('token') !== token) return null
  return `${TOKEN_COOKIE}=${token}; Path=/; HttpOnly; SameSite=Strict`
}

function unauthorized(res: ServerResponse): void {
  res.writeHead(401, { 'Content-Type': 'text/html; charset=utf-8' })
  res.end(
    '<!doctype html><meta charset="utf-8"><title>GenOffice Web</title>' +
      '<p style="font-family:sans-serif">This GenOffice web service requires its access token. ' +
      'Open the URL the app printed at startup (it includes <code>?token=…</code>).</p>',
  )
}

// ---- launcher ----

function launcherHtml(): string {
  const apps: [string, string][] = [
    ['docs', 'Docs'],
    ['sheets', 'Sheets'],
    ['slides', 'Slides'],
    ['pdf', 'PDF'],
    ['markdown', 'Markdown'],
    ['html', 'HTML'],
  ]
  const links = apps
    .map(([id, label]) => `<a class="card" href="/${id}/"><b>${label}</b><span>/${id}/</span></a>`)
    .join('')
  return `<!doctype html><html><head><meta charset="utf-8"><title>GenOffice Web</title>
<style>
body{font-family:-apple-system,system-ui,sans-serif;max-width:720px;margin:48px auto;padding:0 24px;color:#1f2328}
h1{font-size:20px}p{color:#59636e;font-size:14px}
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:24px}
.card{display:flex;flex-direction:column;gap:4px;padding:16px;border:1px solid #d1d9e0;border-radius:10px;text-decoration:none;color:inherit}
.card span{color:#59636e;font-size:12px}
.card:hover{border-color:#0969da}
</style></head><body>
<h1>GenOffice Web</h1>
<p>The desktop suite's engines, editing in a browser tab. Open a document from an
app's start screen; files are read and written by the local service.</p>
<div class="grid">${links}</div>
</body></html>`
}

/** one request: static bundles, media routes, launcher; returns false when the
 *  path belongs to something else (the WebSocket upgrade) */
export async function serveWebStatic(
  req: IncomingMessage,
  res: ServerResponse,
  options: { roots: RendererRoots; token: string | null },
): Promise<boolean> {
  if (options.token && !tokenValid(req, options.token)) {
    unauthorized(res)
    return true
  }
  const setCookie = options.token ? tokenCookieHeader(req, options.token) : null
  if (setCookie) res.setHeader('Set-Cookie', setCookie)

  const url = new URL(req.url ?? '/', 'http://localhost')
  const pathname = decodeURIComponent(url.pathname)

  if (pathname === '/' || pathname === '') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' })
    res.end(launcherHtml())
    return true
  }

  const media = await serveDocsMedia(res, pathname)
  if (media) return true
  if (await serveMarkdownMedia(res, pathname)) return true
  if (await serveHtmlMedia(res, pathname)) return true
  if (servePreview(res, pathname)) return true

  if (pathname.startsWith('/__web/')) {
    const name = pathname.slice('/__web/'.length)
    if (!/^[\w-]+\.js$/.test(name)) {
      res.writeHead(404)
      res.end()
      return true
    }
    const file = join(webBundleDir(), name)
    if (!existsSync(file)) {
      res.writeHead(404, { 'Content-Type': 'text/plain' })
      res.end('web client bundle missing; run tools/build-web-preloads.mjs')
      return true
    }
    res.writeHead(200, {
      'Content-Type': 'text/javascript; charset=utf-8',
      'Cache-Control': 'no-store',
    })
    res.end(readFileSync(file))
    return true
  }

  const mount = /^\/([a-z]+)(\/.*)?$/.exec(pathname)
  if (mount) {
    const appName = mount[1]
    // a known app whose renderer is missing still answers with the build hint
    const root = options.roots[appName] ?? join(app.getAppPath(), '..', appName, 'out', 'renderer')
    const rest = mount[2] ?? '/'
    if (rest === '/' || rest === '') {
      const index = join(root, 'index.html')
      if (existsSync(index)) {
        serveFile(res, index, appName)
        return true
      }
      res.writeHead(404, { 'Content-Type': 'text/plain' })
      res.end(`renderer for "${appName}" is not built; run npm run build -w @genoffice/${appName}`)
      return true
    }
    const schemeUrl = `http://${appName}${rest}`
    const file = resolveRendererFile(new Map(Object.entries(options.roots)), schemeUrl)
    if (file && existsSync(file) && statSync(file).isFile()) {
      serveFile(res, file, appName)
      return true
    }
  }

  res.writeHead(404, { 'Content-Type': 'text/plain' })
  res.end('not found')
  return true
}
