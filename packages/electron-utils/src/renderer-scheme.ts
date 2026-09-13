import type { CustomScheme } from 'electron'
import { resolve, sep } from 'node:path'

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

export type RendererHost = 'docs' | 'sheets' | 'slides' | 'pdf' | 'markdown' | 'html'

/** Dev server URL when one is configured, otherwise the module's scheme URL; the
 * query is appended either way so a dev URL that already carries params stays valid. */
export function rendererUrl(
  devUrl: string | undefined,
  host: RendererHost,
  query?: Record<string, string>,
): string {
  const url = new URL(devUrl ?? `${RENDERER_SCHEME}://${host}/index.html`)
  for (const [key, value] of Object.entries(query ?? {})) url.searchParams.set(key, value)
  return url.toString()
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
  const file = resolve(root, pathname.replace(/^\/+/, ''))
  return file === root || file.startsWith(root + sep) ? file : null
}
