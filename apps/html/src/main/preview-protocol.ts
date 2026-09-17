import { DOCX_MEDIA_SCHEME_PRIVILEGE, RENDERER_SCHEME_PRIVILEGE } from '@genoffice/electron-utils'
import { protocol } from 'electron'
import { ASSET_SCHEME, PREVIEW_SCHEME, buildPreviewDocument } from './preview-document'

export { assetBaseHref, previewUrlFor } from './preview-document'

/** Must run before app ready, and only once per process, so the renderer scheme
 * rides along. Both html schemes are secure: a secure preview document loading
 * stylesheets and scripts from a non-secure scheme would be blocked as mixed content. */
export function registerPrivilegedSchemes(): void {
  protocol.registerSchemesAsPrivileged([
    RENDERER_SCHEME_PRIVILEGE,
    DOCX_MEDIA_SCHEME_PRIVILEGE,
    {
      scheme: PREVIEW_SCHEME,
      privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true },
    },
    {
      scheme: ASSET_SCHEME,
      privileges: {
        standard: true,
        secure: true,
        supportFetchAPI: true,
        corsEnabled: true,
        stream: true,
      },
    },
  ])
}

export function registerPreviewProtocol(
  resolve: (webContentsId: number) => { text: string; baseHref: string | null } | null,
): void {
  protocol.handle(PREVIEW_SCHEME, (request) => {
    let host: string
    try {
      host = new URL(request.url).host
    } catch {
      return new Response(null, { status: 400 })
    }
    const match = /^view-(\d+)$/.exec(host)
    const entry = match ? resolve(Number(match[1])) : null
    if (!entry) return new Response(null, { status: 404 })
    return new Response(buildPreviewDocument(entry.text, entry.baseHref), {
      headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' },
    })
  })
}
