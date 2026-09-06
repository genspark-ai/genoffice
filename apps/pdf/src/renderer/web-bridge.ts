/// Web-version bootstrap for the pdf renderer.
///
/// Loaded before main.tsx and only acts outside Electron: it installs the same
/// `window.pdfApi` / `window.projectApi` objects the preload exposes in the
/// desktop app, but backed by the HTTP/SSE transport against the running
/// Electron main process. Inside Electron the preload has already exposed the
/// IPC-backed APIs and this module leaves them untouched.
import { createHttpIpcTransport, isElectronRuntime } from '@genoffice/ipc-bridge/client'
import { createPdfApi, createPdfProjectApi } from '../shared/pdf-api-factory'

if (!isElectronRuntime()) {
  const transport = createHttpIpcTransport()
  const bridgedWindow = window as unknown as Record<string, unknown>
  // `#open=<path>` grants the path to the bridge sender and returns it as the
  // pending open, so the renderer's boot consumePending() opens it (the web
  // equivalent of the shell granting a path to a real PDF view).
  const hashOpen = new URLSearchParams(window.location.hash.slice(1)).get('open')
  bridgedWindow.pdfApi = createPdfApi(transport, {
    consumePending: async () => {
      if (hashOpen) {
        const granted: unknown = await transport.invoke('pdf:open-path', hashOpen)
        return typeof granted === 'string' && granted ? granted : null
      }
      return null
    },
  })
  bridgedWindow.projectApi = createPdfProjectApi(transport)
}
