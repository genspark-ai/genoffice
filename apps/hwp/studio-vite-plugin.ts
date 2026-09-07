import { cpSync, createReadStream, existsSync, statSync } from 'node:fs'
import { extname, join, resolve } from 'node:path'
import type { IncomingMessage, ServerResponse } from 'node:http'
import type { Plugin } from 'vite'

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

export const STUDIO_MOUNT = '/rhwp'
export const STUDIO_DIR = resolve(__dirname, 'vendor/rhwp-studio')

function studioFile(urlPath: string): string | null {
  const rel = urlPath.split('?')[0].slice(STUDIO_MOUNT.length)
  const base = (rel === '' || rel === '/' ? 'index.html' : rel.replace(/^\//, '')).split('/').pop()
  if (base === 'sw.js' || base === 'registerSW.js' || base === 'manifest.webmanifest') return null
  const wanted = resolve(STUDIO_DIR, rel === '' || rel === '/' ? 'index.html' : `.${rel}`)
  if (!wanted.startsWith(STUDIO_DIR)) return null
  if (existsSync(wanted) && statSync(wanted).isFile()) return wanted
  const fallback = join(STUDIO_DIR, 'index.html')
  return existsSync(fallback) ? fallback : null
}

function middleware(req: IncomingMessage, res: ServerResponse, next: () => void): void {
  const url = req.url ?? ''
  if (!url.startsWith(STUDIO_MOUNT)) {
    next()
    return
  }
  const file = studioFile(url)
  if (!file) {
    res.statusCode = 404
    res.end('rhwp-studio is not vendored — run npm run vendor:studio -w @genoffice/hwp')
    return
  }
  res.setHeader('Content-Type', MIME[extname(file)] ?? 'application/octet-stream')
  createReadStream(file).pipe(res)
}

/** Serve the vendored studio at /rhwp/ (matches the published --base=/rhwp/ build). */
export function rhwpStudioPlugin(): Plugin {
  return {
    name: 'rhwp-studio-static',
    configureServer(server) {
      server.middlewares.use(middleware)
    },
    configurePreviewServer(server) {
      server.middlewares.use(middleware)
    },
    closeBundle() {
      if (!existsSync(join(STUDIO_DIR, 'index.html'))) {
        throw new Error('rhwp-studio is not vendored — run npm run vendor:studio -w @genoffice/hwp')
      }
      const dest = resolve(__dirname, 'out/renderer/rhwp')
      cpSync(STUDIO_DIR, dest, { recursive: true })
    },
  }
}
