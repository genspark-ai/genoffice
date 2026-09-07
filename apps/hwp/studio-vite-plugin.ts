import { cpSync, createReadStream, existsSync } from 'node:fs'
import { join, resolve } from 'node:path'
import type { IncomingMessage, ServerResponse } from 'node:http'
import type { Plugin } from 'vite'
import { existingFile, isBlockedStudioAsset, mimeFor, safePathUnder } from './src/shared/static-serve'

export const STUDIO_MOUNT = '/rhwp'
export const STUDIO_DIR = resolve(__dirname, 'vendor/rhwp-studio')

function studioFile(urlPath: string): string | null {
  if (isBlockedStudioAsset(urlPath)) return null
  const rel = urlPath.split('?')[0].slice(STUDIO_MOUNT.length)
  const wanted = safePathUnder(STUDIO_DIR, rel === '' || rel === '/' ? '/index.html' : rel)
  if (!wanted) return null
  return existingFile(wanted) ?? existingFile(join(STUDIO_DIR, 'index.html'))
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
  res.setHeader('Content-Type', mimeFor(file))
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
