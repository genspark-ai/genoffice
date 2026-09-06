/// Web-version bootstrap for the sheets renderer.
///
/// Loaded before main.tsx and only acts outside Electron: it installs the same
/// `window.desktopApi` / `window.projectApi` objects the preload exposes in
/// the desktop app, but backed by the HTTP/SSE transport against the running
/// Electron main process. Native-only channels (workbook open dialogs, CSV
/// save confirm, attachment picker, screen capture) get browser equivalents so
/// the web version keeps the full feature surface. Inside Electron the preload
/// has already exposed the IPC-backed APIs and this module leaves them
/// untouched.
import { createHttpIpcTransport, isElectronRuntime } from '@genoffice/ipc-bridge/client'
import { createWebFileBridge, pickFileBytes } from '@genoffice/ipc-bridge/web-native'
import { createSheetsApi, createSheetsProjectApi } from '../shared/sheets-api-factory'

if (!isElectronRuntime()) {
  const transport = createHttpIpcTransport()
  const files = createWebFileBridge(transport)
  const bridgedWindow = window as unknown as Record<string, unknown>
  bridgedWindow.desktopApi = createSheetsApi(transport, {
    selectWorkbook: async () => {
      const picked = await pickFileBytes('.xlsx,.xlsm,.xls,.csv')
      if (!picked) return null
      const file = picked[0]
      if (!file) return null
      const { name, bytes } = file
      const path = await files.writeTempFile(name, bytes)
      return await transport.invoke('workbook:open-path', path)
    },
    selectWorkbooksForMerge: async () => {
      const picked = await pickFileBytes('.xlsx,.xlsm,.xls,.csv', true)
      if (!picked) return null
      const paths: string[] = []
      for (const file of picked) {
        paths.push(await files.writeTempFile(file.name, file.bytes))
      }
      return await transport.invoke('workbook:open-for-merge', paths)
    },
    confirmCsvSave: async () => 'csv',
    pickAttachments: async () => {
      const picked = await pickFileBytes(undefined, true)
      if (!picked) return null
      const paths: string[] = []
      for (const file of picked) {
        paths.push(await files.writeTempFile(file.name, file.bytes))
      }
      return await transport.invoke('sheets:files-add', paths)
    },
    captureScreenSources: async () => {
      return { status: 'ok', sources: [{ id: 'display', name: 'Screen', kind: 'screen', thumbnail: '' }] }
    },
    captureScreenSource: async () => {
      return await captureDisplayFrame()
    },
  })
  bridgedWindow.projectApi = createSheetsProjectApi(transport)
}

async function captureDisplayFrame(): Promise<{
  mediaType: 'image/png'
  base64: string
  width: number
  height: number
} | null> {
  try {
    const stream = await navigator.mediaDevices.getDisplayMedia({ video: true })
    const video = document.createElement('video')
    video.srcObject = stream
    await new Promise<void>((resolve) => {
      video.onloadedmetadata = () => resolve()
      video.play()
    })
    const width = video.videoWidth
    const height = video.videoHeight
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const context = canvas.getContext('2d')
    if (!context) {
      stream.getTracks().forEach((track) => track.stop())
      return null
    }
    context.drawImage(video, 0, 0, width, height)
    stream.getTracks().forEach((track) => track.stop())
    return { mediaType: 'image/png', base64: canvas.toDataURL('image/png').split(',')[1] ?? '', width, height }
  } catch {
    return null
  }
}
