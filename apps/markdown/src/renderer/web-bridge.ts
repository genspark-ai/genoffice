/// Web-version bootstrap for the markdown renderer.
///
/// Loaded before main.tsx and only acts outside Electron: it installs the same
/// `window.markdownApi` / `window.projectApi` objects the preload exposes in
/// the desktop app, but backed by the HTTP/SSE transport against the running
/// Electron main process. Native-only channels (image picker, DOCX/PDF export)
/// get browser equivalents so the web version keeps the full feature surface.
/// Inside Electron the preload has already exposed the IPC-backed APIs and
/// this module leaves them untouched.
import { createHttpIpcTransport, isElectronRuntime } from '@genoffice/ipc-bridge/client'
import { downloadBytes, pickFileBytes } from '@genoffice/ipc-bridge/web-native'
import { createMarkdownApi, createMarkdownProjectApi } from '../shared/markdown-api-factory'

if (!isElectronRuntime()) {
  const transport = createHttpIpcTransport()
  const bridgedWindow = window as unknown as Record<string, unknown>
  bridgedWindow.markdownApi = createMarkdownApi(transport, {
    pickImage: async () => {
      const picked = await pickFileBytes('image/png,image/jpeg,image/gif')
      if (!picked) return null
      const { name, bytes } = picked[0]
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      if (!['png', 'jpg', 'jpeg', 'gif'].includes(ext)) return null
      const base64 = bytesToBase64(bytes)
      return await transport.invoke('markdown:save-image', { base64, ext })
    },
    exportDocx: async (request) => {
      if (typeof request?.base64 !== 'string' || !request.base64) {
        return { ok: false, error: 'markdown: bad export request' }
      }
      const safeName =
        String(request.suggestedName || 'Untitled').replace(/[/\\:*?"<>|]/g, '_') || 'Untitled'
      downloadBytes(`${safeName}.docx`, base64ToBytes(request.base64))
      return { ok: true, path: '' }
    },
    exportPdf: async (request) => {
      if (typeof request?.html !== 'string' || !request.html) {
        return { ok: false, error: 'markdown: bad export request' }
      }
      const win = window.open('', '_blank')
      if (!win) return { ok: false, error: 'web: popup blocked' }
      win.document.open()
      win.document.write(request.html)
      win.document.close()
      win.addEventListener('load', () => win.print())
      return { ok: true, path: '' }
    },
  })
  bridgedWindow.projectApi = createMarkdownProjectApi(transport)
}

function bytesToBase64(bytes: ArrayBuffer): string {
  let binary = ''
  const view = new Uint8Array(bytes)
  const CHUNK = 0x8000
  for (let i = 0; i < view.length; i += CHUNK) {
    binary += String.fromCharCode(...view.subarray(i, i + CHUNK))
  }
  return btoa(binary)
}

function base64ToBytes(base64: string): ArrayBuffer {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes.buffer
}
