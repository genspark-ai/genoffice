/// Web-version bootstrap for the docs renderer.
///
/// Loaded before main.tsx and only acts outside Electron: it installs the same
/// `window.desktop` / `window.projectApi` objects the preload exposes in the
/// desktop app, but backed by the HTTP/SSE transport against the running
/// Electron main process. Native-only channels (open/save dialogs, print,
/// clipboard, font metrics, window management) get browser equivalents so the
/// web version keeps the full feature surface. Inside Electron the preload has
/// already exposed the IPC-backed APIs and this module leaves them untouched.
import { createHttpIpcTransport, isElectronRuntime } from '@genoffice/ipc-bridge/client'
import {
  createWebFileBridge,
  downloadBytes,
  pickFileBytes,
  webCopyImage,
  webFontMetrics,
  webOpenTab,
  webPrint,
} from '@genoffice/ipc-bridge/web-native'
import { createDesktopApi, createProjectApi } from '../shared/desktop-api-factory'

if (!isElectronRuntime()) {
  const transport = createHttpIpcTransport()
  const files = createWebFileBridge(transport)
  const bridgedWindow = window as unknown as Record<string, unknown>
  bridgedWindow.desktop = createDesktopApi(transport, {
    openDocx: async () => {
      const picked = await pickFileBytes('.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document')
      if (!picked) return null
      const { name, bytes } = picked[0]
      const path = await files.writeTempFile(name, bytes)
      return await transport.invoke('docs:open-path', path)
    },
    pickImage: async () => {
      const picked = await pickFileBytes('image/png,image/jpeg,image/gif')
      if (!picked) return null
      const { name, bytes } = picked[0]
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      const mime = IMAGE_MIME[ext]
      if (!mime) return null
      return { base64: bytesToBase64(bytes), mime, name }
    },
    saveDocxAs: async (defaultName, data) => {
      downloadBytes(defaultName, data)
      const path = await files.writeTempFile(defaultName, data)
      return { ok: true, path }
    },
    saveDocxNew: async (defaultName, data) => {
      downloadBytes(defaultName, data)
      const path = await files.writeTempFile(defaultName, data)
      return { ok: true, path }
    },
    print: async () => {
      webPrint()
      return { ok: true }
    },
    exportPdf: async () => {
      webPrint()
      return { ok: true, path: '' }
    },
    copyImageToClipboard: (dataUrl) => webCopyImage(dataUrl),
    fontMetrics: (family) => Promise.resolve(webFontMetrics(family)),
    pickAttachments: async () => {
      const picked = await pickFileBytes(undefined, true)
      if (!picked) return null
      const paths: string[] = []
      for (const file of picked) {
        paths.push(await files.writeTempFile(file.name, file.bytes))
      }
      return await transport.invoke('files:add', paths)
    },
    openNewTab: async (openPath) => {
      webOpenTab(openPath ? `#open=${encodeURIComponent(openPath)}` : window.location.href)
    },
    listDocsTabs: async () => [],
    focusDocsTab: async () => {},
    printPdfBuffer: async () => {
      // The browser cannot produce PDF bytes per print group; the merged save
      // below opens the browser print dialog (save as PDF) instead.
      return { ok: true, base64: '' }
    },
    saveMergedPdf: async () => {
      webPrint()
      return { ok: true, path: '' }
    },
  })
  bridgedWindow.projectApi = createProjectApi(transport)
}

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
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
