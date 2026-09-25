import { basename } from 'node:path'
import { net, protocol } from 'electron'
import { pathToFileURL } from 'node:url'
import {
  DOCX_MEDIA_SCHEME_PRIVILEGE,
  RENDERER_SCHEME,
  RENDERER_SCHEME_PRIVILEGE,
  SCRIPT_WORKER_CSP,
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
  protocol.handle(RENDERER_SCHEME, async (request) => {
    const file = resolveRendererFile(table, request.url)
    if (!file) return new Response(null, { status: 404 })
    const res = await net
      .fetch(pathToFileURL(file).toString())
      .catch(() => new Response(null, { status: 404 }))
    // Script-sandbox worker chunks must carry the sandbox CSP on their own response
    // (a dedicated worker never sees the page's CSP); see SCRIPT_WORKER_CSP.
    if (/^script-worker-.*\.js$/.test(basename(file))) {
      return new Response(await res.arrayBuffer(), {
        status: res.status,
        headers: {
          'content-type': 'text/javascript; charset=utf-8',
          'content-security-policy': SCRIPT_WORKER_CSP,
        },
      })
    }
    return res
  })
}
