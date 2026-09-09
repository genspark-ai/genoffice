/// Local store for AI-generated images that arrive as bytes (BYOK image
/// providers answer with base64, unlike the Genspark CDN URLs the insert
/// pipelines were built around). The bytes are written to a fixed temp
/// directory and handed back as a file:// URL, which is the only file:// shape
/// fetchRemoteImage accepts — anything outside this directory stays refused.

import { randomUUID } from 'node:crypto'
import { mkdirSync, readFileSync, readdirSync, statSync, unlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join, resolve, sep } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

export const GENERATED_IMAGE_DIR = join(tmpdir(), 'genoffice-ai-images')

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
}
const EXT_BY_MIME: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/jpg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
}
const FILE_NAME_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.(png|jpg|gif|webp)$/

const MAX_AGE_MS = 24 * 60 * 60 * 1000

/** Best-effort sweep of yesterday's images; the directory is a cache, not a document store */
function pruneOld(now = Date.now()): void {
  let names: string[]
  try {
    names = readdirSync(GENERATED_IMAGE_DIR)
  } catch {
    return
  }
  for (const name of names) {
    if (!FILE_NAME_RE.test(name)) continue
    const path = join(GENERATED_IMAGE_DIR, name)
    try {
      if (now - statSync(path).mtimeMs > MAX_AGE_MS) unlinkSync(path)
    } catch {
      /* another process may have removed it */
    }
  }
}

export function storeGeneratedImage(bytes: Uint8Array, mime: string): string {
  const ext = EXT_BY_MIME[mime.toLowerCase()] ?? 'png'
  mkdirSync(GENERATED_IMAGE_DIR, { recursive: true })
  pruneOld()
  const path = join(GENERATED_IMAGE_DIR, `${randomUUID()}.${ext}`)
  writeFileSync(path, bytes)
  return pathToFileURL(path).toString()
}

/**
 * Reads a file:// URL back only when it names an image this store wrote:
 * inside GENERATED_IMAGE_DIR, uuid-named, image extension. Returns null for
 * every other URL so a prompt-injected file:///etc/passwd never gets read.
 */
export function readGeneratedImage(rawUrl: string): { bytes: Buffer; mime: string } | null {
  let path: string
  try {
    const url = new URL(rawUrl)
    if (url.protocol !== 'file:') return null
    path = resolve(fileURLToPath(url))
  } catch {
    return null
  }
  const dir = resolve(GENERATED_IMAGE_DIR)
  if (!path.startsWith(dir + sep)) return null
  const name = basename(path)
  if (!FILE_NAME_RE.test(name) || path !== join(dir, name)) return null
  try {
    const bytes = readFileSync(path)
    return { bytes, mime: MIME_BY_EXT[name.slice(name.lastIndexOf('.') + 1)] ?? 'image/png' }
  } catch {
    return null
  }
}
