/// "Save Image As…" for any image the renderer displays: data URLs are decoded
/// in place, everything else (http(s), app asset schemes) goes through
/// net.fetch so custom protocol handlers keep enforcing their own access rules.
import { writeFile } from 'node:fs/promises'
import type { BrowserWindow } from 'electron'
import { showSaveDialogWithMemory } from './dialog-memory'
import { MAX_REMOTE_IMAGE_BYTES, readBodyCapped } from './remote-image'
import { fetchWithSsrfGuard } from './safe-remote-url'

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
  // The parameter run is `;`-prefixed and its body excludes `;`, so each
  // iteration has exactly one possible length. With `;[^,]*` instead, one
  // iteration could swallow several parameters or several could cover one, and
  // a URL with no comma made the engine try every split - exponential, so a
  // ~100-character data: URL in a document froze the main process when the user
  // chose "Save image".
  const m = /^data:([^;,]*)((?:;[^;,]*)*),([\s\S]*)$/i.exec(url)
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
  // Remote http(s) URLs go through the SSRF guard, which validates every
  // redirect hop against internal addresses — this was the only outbound fetch
  // in the package that bypassed it, so a compromised renderer could make the
  // main process fetch internal URLs and write the response to a user-chosen
  // file. App asset schemes (md-asset:// and siblings) are process-local
  // protocol handlers, not network addresses; they fetch directly.
  const res = /^https?:/i.test(url)
    ? await fetchWithSsrfGuard(url, {
        // Electron's net.fetch types a narrower input than the DOM fetch; both
        // take a string URL here
        fetchImpl: net.fetch as unknown as typeof fetch,
      })
    : await net.fetch(url)
  if (!res) throw new Error('image url blocked (internal or unsafe address)')
  if (!res.ok) throw new Error(`fetch failed: HTTP ${res.status}`)
  return {
    bytes: Buffer.from(await readBodyCapped(res, MAX_REMOTE_IMAGE_BYTES)),
    mime: res.headers.get('content-type'),
  }
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
