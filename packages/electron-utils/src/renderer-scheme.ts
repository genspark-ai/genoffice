import type { CustomScheme } from 'electron'
import { resolve, sep } from 'node:path'

import { GENERATED_IMAGE_DIR } from './generated-images'
import { GENERATED_IMAGE_FILE_RE, GENERATED_IMAGE_URL_PREFIX } from './generated-image-url'

/** Origin the built renderers are served from (`genoffice-app://<module>/index.html`). */
export const RENDERER_SCHEME = 'genoffice-app'

/** Chromium persists V8 code caches only for http(s) and for privileged custom
 * schemes; a file:// renderer recompiled its whole bundle on every open. */
export const RENDERER_SCHEME_PRIVILEGE: CustomScheme = {
  scheme: RENDERER_SCHEME,
  privileges: {
    standard: true,
    secure: true,
    supportFetchAPI: true,
    corsEnabled: true,
    stream: true,
    codeCache: true,
  },
}

/** docs' lazily served pictures (packages/docx-engine/src/lazy-media.ts); a
 * secure page may only load pictures from another secure scheme */
export const DOCX_MEDIA_SCHEME_PRIVILEGE: CustomScheme = {
  scheme: 'genoffice-docx-media',
  privileges: {
    standard: true,
    secure: true,
    supportFetchAPI: true,
    corsEnabled: true,
    stream: true,
  },
}

/** Slides serves archive pictures outside structured-clone IPC payloads. */
export const SLIDE_MEDIA_SCHEME_PRIVILEGE: CustomScheme = {
  scheme: 'genoffice-slide-media',
  privileges: {
    standard: true,
    secure: true,
    supportFetchAPI: true,
    corsEnabled: true,
    stream: true,
  },
}

export type RendererHost = 'docs' | 'sheets' | 'slides' | 'pdf' | 'markdown' | 'html' | 'help'

export const MAX_RENDERER_QUERY_ENTRIES = 20
export const MAX_RENDERER_QUERY_CHARS = 4_000

/** Dev server URL when one is configured, otherwise the module's scheme URL; the
 * query is appended either way so a dev URL that already carries params stays valid. */
export function rendererUrl(
  devUrl: string | undefined,
  host: RendererHost,
  query?: Record<string, string>,
): string {
  let url: URL
  try {
    url = new URL(devUrl ?? `${RENDERER_SCHEME}://${host}/index.html`)
  } catch {
    throw new Error(`Invalid dev URL for renderer "${host}": "${devUrl}"`)
  }
  const entries = Object.entries(query ?? {})
  if (entries.length > MAX_RENDERER_QUERY_ENTRIES) {
    throw new Error(
      `Too many renderer query params (${entries.length}, cap ${MAX_RENDERER_QUERY_ENTRIES})`,
    )
  }
  for (const [key, value] of entries) {
    if (key.length > 256 || value.length > MAX_RENDERER_QUERY_CHARS) {
      throw new Error('Renderer query param too long')
    }
    url.searchParams.set(key, value)
  }
  return url.toString()
}

/**
 * Same-origin path for a BYOK-generated image preview: a secure renderer page
 * cannot load the store's file:// URL, so the protocol serves that one folder
 * under `/generated-images/<name>` instead. Only the store's own UUID-shaped
 * names are accepted, so the route cannot reach any other file.
 */
export function resolveGeneratedImageFile(requestUrl: string): string | null {
  let url: URL
  try {
    url = new URL(requestUrl)
  } catch {
    return null
  }
  if (url.protocol !== `${RENDERER_SCHEME}:`) return null
  if (!url.pathname.startsWith(GENERATED_IMAGE_URL_PREFIX)) return null
  let name: string
  try {
    name = decodeURIComponent(url.pathname.slice(GENERATED_IMAGE_URL_PREFIX.length))
  } catch {
    return null
  }
  if (!GENERATED_IMAGE_FILE_RE.test(name)) return null
  const file = resolve(GENERATED_IMAGE_DIR, name)
  return file.startsWith(GENERATED_IMAGE_DIR + sep) ? file : null
}

/** Map a scheme request onto a file under the host's renderer directory; null
 * for unknown hosts and for paths that escape the directory. */
export function resolveRendererFile(
  roots: ReadonlyMap<string, string>,
  requestUrl: string,
): string | null {
  let url: URL
  try {
    url = new URL(requestUrl)
  } catch {
    return null
  }
  const root = roots.get(url.hostname)
  if (!root) return null
  let pathname: string
  try {
    pathname = decodeURIComponent(url.pathname)
  } catch {
    return null
  }
  if (pathname.length > 4096 || pathname.includes('\0')) return null
  // Treat backslashes as separators too: %5c decodes to `\`, which resolves
  // as a directory separator on Windows but not on POSIX.
  const relative = pathname.replace(/\\/g, '/').replace(/^\/+/, '')
  // Normalize the root first so trailing slashes or mixed separators cannot
  // break the containment check below.
  const base = resolve(root)
  const file = resolve(base, relative)
  // The root itself is a directory, not a file to serve.
  if (file === base) return null
  return file.startsWith(base + sep) ? file : null
}
