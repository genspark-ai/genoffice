import { net, protocol } from 'electron'
import { pathToFileURL } from 'node:url'
import {
  DOCX_MEDIA_SCHEME_PRIVILEGE,
  RENDERER_SCHEME,
  RENDERER_SCHEME_PRIVILEGE,
  resolveRendererFile,
} from './renderer-scheme'

/** Before app ready; a process may call registerSchemesAsPrivileged only once,
 * so hosts that register other schemes spread RENDERER_SCHEME_PRIVILEGE into
 * their own list instead. */
export function registerRendererScheme(): void {
  protocol.registerSchemesAsPrivileged([RENDERER_SCHEME_PRIVILEGE, DOCX_MEDIA_SCHEME_PRIVILEGE])
}

/** After app ready: serve each module's built renderer directory. */
export function installRendererProtocol(roots: Record<string, string>): void {
  const table = new Map(Object.entries(roots))
  protocol.handle(RENDERER_SCHEME, (request) => {
    const file = resolveRendererFile(table, request.url)
    if (!file) return new Response(null, { status: 404 })
    return net
      .fetch(pathToFileURL(file).toString())
      .catch(() => new Response(null, { status: 404 }))
  })
}
