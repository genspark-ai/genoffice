/**
 * Serving for the sheets scripting sandbox worker (issue #815).
 *
 * A dedicated worker takes its Content-Security-Policy from its own response —
 * the page's CSP never reaches it, which is why the page's `script-src 'self'`
 * does not stop a worker from fetching file:// or arbitrary http(s). The worker
 * script is therefore served from its own genoffice-script:// protocol whose
 * handler attaches the sandbox CSP: eval stays available (user scripts compile
 * with `new Function`), while every network request and cross-origin script
 * load — including dynamic import() — is denied at the engine level.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { net, protocol } from 'electron'
import { SCRIPT_SCHEME } from '@genoffice/electron-utils/renderer-scheme'

export const SCRIPT_WORKER_URL = `${SCRIPT_SCHEME}://worker/script-worker.js`

const SCRIPT_WORKER_CSP = "default-src 'none'; script-src 'self' 'unsafe-eval'; connect-src 'none'"

let cachedBundledSource: string | null = null

/** The worker source: proxied from the dev server in dev, the built chunk in production. */
async function workerSource(
  devUrl: string | undefined,
  rendererDir: string,
): Promise<string | null> {
  if (devUrl) {
    const base = devUrl.endsWith('/') ? devUrl : `${devUrl}/`
    const res = await net.fetch(new URL('src/scripting/script-worker.ts', base).toString())
    if (!res.ok) return null
    return res.text()
  }
  if (cachedBundledSource) return cachedBundledSource
  const assets = join(rendererDir, 'assets')
  const names = readdirSync(assets).filter((n) => /^script-worker-.*\.js$/.test(n))
  if (names.length !== 1) return null
  cachedBundledSource = readFileSync(join(assets, names[0]!), 'utf8')
  return cachedBundledSource
}

/** After app ready: serve the sandbox worker with the sandbox CSP attached. */
export function installScriptWorkerProtocol(devUrl: string | undefined, rendererDir: string): void {
  protocol.handle(SCRIPT_SCHEME, async (request) => {
    let path: string
    try {
      path = new URL(request.url).pathname
    } catch {
      return new Response(null, { status: 404 })
    }
    if (path !== '/script-worker.js') return new Response(null, { status: 404 })
    try {
      const source = await workerSource(devUrl, rendererDir)
      if (source == null) return new Response(null, { status: 404 })
      return new Response(source, {
        headers: {
          'content-type': 'text/javascript; charset=utf-8',
          'content-security-policy': SCRIPT_WORKER_CSP,
        },
      })
    } catch {
      return new Response(null, { status: 404 })
    }
  })
}
