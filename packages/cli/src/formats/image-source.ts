import { existsSync, readFileSync } from 'node:fs'
import { extname, isAbsolute, resolve } from 'node:path'
import {
  MAX_REMOTE_IMAGE_BYTES,
  ResponseTooLargeError,
  fetchRemoteImage,
  readBodyCapped,
} from '@genoffice/electron-utils/remote-image'
import { assertAllowed, type PathContext } from '../fs'
import { imageSize } from './image-size'

export interface ImageSource {
  bytes: Uint8Array
  ext: string
  mime: string | null
}

/**
 * Image bytes for a spec or op field: a data: URL, an http(s) URL genoffice
 * downloads, or a local path (absolute, else relative to the current directory
 * and then to the file the reference came from).
 */
export async function readImageSource(
  url: string,
  ctx: PathContext,
  baseDir?: string,
): Promise<ImageSource | null> {
  if (url.startsWith('data:')) {
    const m = /^data:image\/([a-z0-9.+-]+);base64,(.*)$/is.exec(url)
    if (!m) return null
    const bytes = new Uint8Array(Buffer.from(m[2]!, 'base64'))
    return withType(bytes, m[1]!.toLowerCase())
  }
  if (/^https?:\/\//i.test(url)) {
    const resp = await fetchRemoteImage(url)
    if (!resp || !resp.ok) return null
    const bytes = await readBodyCapped(resp, MAX_REMOTE_IMAGE_BYTES).catch((err: unknown) => {
      if (err instanceof ResponseTooLargeError) return null
      throw err
    })
    if (!bytes) return null
    const declared = (resp.headers.get('content-type') ?? '').split('/')[1]?.split(';')[0] ?? ''
    return withType(bytes, declared)
  }
  const candidates = isAbsolute(url)
    ? [url]
    : [resolve(ctx.cwd, url), ...(baseDir ? [resolve(baseDir, url)] : [])]
  const path = candidates.find((p) => existsSync(p))
  if (!path) return null
  assertAllowed(path, ctx.env, 'read')
  const bytes = new Uint8Array(readFileSync(path))
  return withType(bytes, extname(path).slice(1).toLowerCase())
}

/** Sniffed type wins over what the name or header claims (a .jpg holding PNG bytes is common). */
function withType(bytes: Uint8Array, declared: string): ImageSource {
  const mime = imageSize(bytes)?.mime ?? null
  const ext =
    mime === 'image/png'
      ? 'png'
      : mime === 'image/jpeg'
        ? 'jpg'
        : mime === 'image/gif'
          ? 'gif'
          : declared === 'jpeg'
            ? 'jpg'
            : declared || 'png'
  return { bytes, ext, mime }
}
