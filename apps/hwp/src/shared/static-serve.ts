import { existsSync, statSync } from 'node:fs'
import { extname, resolve, sep } from 'node:path'

export const STATIC_MIME: Record<string, string> = {
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

const BLOCKED_STUDIO = new Set(['sw.js', 'registerSW.js', 'manifest.webmanifest'])

export function mimeFor(file: string): string {
  return STATIC_MIME[extname(file)] ?? 'application/octet-stream'
}

export function isBlockedStudioAsset(urlPath: string): boolean {
  const base = urlPath.split('?')[0].split('/').pop() ?? ''
  return BLOCKED_STUDIO.has(base)
}

function decodePath(urlPath: string): string | null {
  try {
    return decodeURIComponent(urlPath.split('?')[0])
  } catch {
    return null
  }
}

/** Resolved path under `root`, or null on traversal / blocked / malformed. */
export function safePathUnder(root: string, urlPath: string): string | null {
  if (isBlockedStudioAsset(urlPath)) return null
  const decoded = decodePath(urlPath)
  if (decoded === null) return null
  const rel = decoded === '/' || decoded === '' ? '/index.html' : decoded
  const normalized = rel.startsWith('/') ? rel : `/${rel}`
  const base = resolve(root)
  const wanted = resolve(base, `.${normalized}`)
  if (wanted !== base && !wanted.startsWith(base + sep)) return null
  return wanted
}

export function existingFile(path: string | null): string | null {
  if (!path) return null
  return existsSync(path) && statSync(path).isFile() ? path : null
}
