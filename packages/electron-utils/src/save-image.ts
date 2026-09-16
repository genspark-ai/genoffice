/// "Save Image As…" for any image the renderer displays: data URLs are decoded
/// in place, everything else (http(s), app asset schemes) goes through
/// net.fetch so custom protocol handlers keep enforcing their own access rules.
import { writeFile } from 'node:fs/promises'
import type { BrowserWindow } from 'electron'
import { showSaveDialogWithMemory } from './dialog-memory'

const EXT_BY_MIME: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
  'image/bmp': 'bmp',
  'image/svg+xml': 'svg',
  'image/avif': 'avif',
  'image/tiff': 'tif',
}
const IMAGE_FILE = /\.(png|jpe?g|gif|webp|bmp|svg|avif|tiff?)$/i

export interface SaveImageResult {
  ok: boolean
  path?: string
  error?: string
}

/** file:// would let a page read arbitrary local files; blob: lives in the renderer only */
export function isSavableImageUrl(url: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(url) && !/^(file|blob|javascript|about):/i.test(url)
}

function urlFileName(url: string): string {
  try {
    return decodeURIComponent(new URL(url).pathname).split('/').pop() ?? ''
  } catch {
    return ''
  }
}

/** The URL's own file name when it carries an image extension, else image.<ext from the MIME type> */
export function suggestImageFileName(url: string, mime: string | null | undefined): string {
  const name = urlFileName(url)
  if (name && IMAGE_FILE.test(name)) return name
  const key = (mime ?? '').split(';')[0]!.trim().toLowerCase()
  return `image.${EXT_BY_MIME[key] ?? 'png'}`
}

export function decodeDataUrl(url: string): { bytes: Buffer; mime: string | null } | null {
  const m = /^data:([^;,]*)((?:;[^,]*)*),([\s\S]*)$/i.exec(url)
  if (!m) return null
  const mime = m[1] || null
  const payload = m[3] ?? ''
  const bytes = /;base64/i.test(m[2] ?? '')
    ? Buffer.from(payload, 'base64')
    : Buffer.from(decodeURIComponent(payload), 'utf8')
  return { bytes, mime }
}

async function fetchImageBytes(url: string): Promise<{ bytes: Buffer; mime: string | null }> {
  if (/^data:/i.test(url)) {
    const decoded = decodeDataUrl(url)
    if (!decoded) throw new Error('malformed data URL')
    return decoded
  }
  const { net } = await import('electron')
  const res = await net.fetch(url)
  if (!res.ok) throw new Error(`fetch failed: HTTP ${res.status}`)
  return { bytes: Buffer.from(await res.arrayBuffer()), mime: res.headers.get('content-type') }
}

export async function saveImageFromUrl(
  parent: BrowserWindow | null | undefined,
  url: string,
  opts: { title: string; fallbackDir?: string },
): Promise<SaveImageResult> {
  if (!isSavableImageUrl(url)) return { ok: false, error: 'unsupported image url' }
  try {
    const { bytes, mime } = await fetchImageBytes(url)
    const name = suggestImageFileName(url, mime)
    const ext = name.slice(name.lastIndexOf('.') + 1)
    const { dialog } = await import('electron')
    const picked = await showSaveDialogWithMemory(
      dialog,
      parent,
      {
        title: opts.title,
        defaultPath: name,
        filters: [{ name: ext.toUpperCase(), extensions: [ext] }],
      },
      opts.fallbackDir,
    )
    if (picked.canceled || !picked.filePath) return { ok: false }
    await writeFile(picked.filePath, bytes)
    return { ok: true, path: picked.filePath }
  } catch (err) {
    return { ok: false, error: String(err) }
  }
}
