import { createReadStream, existsSync, statSync } from 'node:fs'
import { createServer } from 'node:http'
import { extname, join, resolve } from 'node:path'

const MIME: Record<string, string> = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.ico': 'image/x-icon',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ttf': 'font/ttf',
  '.wasm': 'application/wasm',
  '.webmanifest': 'application/manifest+json',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
}

/**
 * rhwp-studio's embed SDK rejects file:/opaque origins. Packaged Hangul tabs
 * therefore load the renderer + vendored studio over loopback HTTP.
 */
export function startHwpLoopback(root: string): Promise<string> {
  const base = resolve(root)
  const server = createServer((req, res) => {
    const raw = decodeURIComponent((req.url ?? '/').split('?')[0])
    const rel = raw === '/' ? '/index.html' : raw
    const wanted = resolve(base, `.${rel}`)
    if (!wanted.startsWith(base)) {
      res.statusCode = 403
      res.end()
      return
    }
    const file =
      existsSync(wanted) && statSync(wanted).isFile()
        ? wanted
        : rel.startsWith('/rhwp/')
          ? join(base, 'rhwp', 'index.html')
          : join(base, 'index.html')
    if (!existsSync(file) || !statSync(file).isFile()) {
      res.statusCode = 404
      res.end()
      return
    }
    res.setHeader('Content-Type', MIME[extname(file)] ?? 'application/octet-stream')
    createReadStream(file).pipe(res)
  })
  return new Promise((resolveAddr, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const addr = server.address()
      if (!addr || typeof addr === 'string') {
        reject(new Error('hwp: loopback bind failed'))
        return
      }
      resolveAddr(`http://127.0.0.1:${addr.port}/`)
    })
  })
}
