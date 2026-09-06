/// Web-version bootstrap for the slides renderer.
///
/// Loaded before main.tsx and only acts outside Electron: it installs the same
/// `window.slidesApi` / `window.desktop` / `window.projectApi` objects the
/// preload exposes in the desktop app, but backed by the HTTP/SSE transport
/// against the running Electron main process. Native-only channels (open,
/// insert pickers, export dirs, print, clipboard, fullscreen, font install)
/// get browser equivalents so the web version keeps the full feature surface.
/// Inside Electron the preload has already exposed the IPC-backed APIs and this
/// module leaves them untouched.
import { createHttpIpcTransport, isElectronRuntime } from '@genoffice/ipc-bridge/client'
import {
  createWebFileBridge,
  downloadBytes,
  pickFileBytes,
  webFullscreen,
  webPrint,
} from '@genoffice/ipc-bridge/web-native'
import {
  createSlidesApi,
  createSlidesFilesApi,
  createSlidesProjectApi,
} from '../shared/slides-api-factory'

if (!isElectronRuntime()) {
  const transport = createHttpIpcTransport()
  const files = createWebFileBridge(transport)
  const bridgedWindow = window as unknown as Record<string, unknown>
  bridgedWindow.slidesApi = createSlidesApi(transport, {
    setShowFullScreen: async () => {
      await webFullscreen()
      return null
    },
    openPptx: async (fitWidthPx) => {
      const picked = await pickFileBytes('.pptx,.ppt,application/vnd.openxmlformats-officedocument.presentationml.presentation')
      if (!picked) return null
      const { name, bytes } = picked[0]
      const path = await files.writeTempFile(name, bytes)
      return await transport.invoke('slides:open-path', path, fitWidthPx)
    },
    insertImage: async (slideIndex, fitWidthPx) => {
      const picked = await pickFileBytes('image/*')
      if (!picked) return null
      const { name, bytes } = picked[0]
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      if (!['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'].includes(ext)) {
        return { error: 'unsupported', ext }
      }
      const natural = await imageNaturalSize(bytes)
      const deckSize = await deckSizePx()
      if (!deckSize) return null
      const maxW = deckSize.cx / 2
      const maxH = deckSize.cy / 2
      const scale = Math.min(maxW / natural.width, maxH / natural.height)
      const cx = Math.round(natural.width * scale)
      const cy = Math.round(natural.height * scale)
      return await transport.invoke('slides:add-image-bytes', {
        slideIndex,
        base64: bytesToBase64(bytes),
        ext,
        xPx: Math.round((deckSize.cx - cx) / 2),
        yPx: Math.round((deckSize.cy - cy) / 2),
        wPx: cx,
        hPx: cy,
        fitWidthPx,
        name,
      })
    },
    insertMedia: async (slideIndex, kind, fitWidthPx) => {
      const picked = await pickFileBytes(kind === 'video' ? 'video/*' : 'audio/*')
      if (!picked) return null
      const { name, bytes } = picked[0]
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      return await transport.invoke('slides:add-media-bytes', {
        slideIndex,
        kind,
        base64: bytesToBase64(bytes),
        ext,
        fitWidthPx,
        name,
      })
    },
    insertModel3d: async (slideIndex, fitWidthPx) => {
      const picked = await pickFileBytes('.glb,.gltf')
      if (!picked) return null
      const { name, bytes } = picked[0]
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      const deckSize = await deckSizePx()
      if (!deckSize) return null
      const cy = Math.round(deckSize.cy * 0.5)
      const cx = cy
      return await transport.invoke('slides:apply-txn', {
        ops: [
          {
            op: 'addModel3d',
            target: { slide: slideIndex },
            bytes: new Uint8Array(bytes),
            ext,
            offset: {
              x: Math.round((deckSize.cx - cx) / 2),
              y: Math.round((deckSize.cy - cy) / 2),
              cx,
              cy,
            },
            name,
          },
        ],
      })
    },
    printSlides: async () => {
      webPrint()
      return { ok: true }
    },
    clipboardExternal: async () => {
      // browser clipboard paste is handled by the renderer's own paste events;
      // the native clipboard probe returns nothing usable over HTTP
      return null
    },
    fontInstallLocal: async () => null,
    pickAttachments: async () => {
      const picked = await pickFileBytes(undefined, true)
      if (!picked) return null
      const paths: string[] = []
      for (const file of picked) {
        paths.push(await files.writeTempFile(file.name, file.bytes))
      }
      return await transport.invoke('slides:files-add', paths)
    },
    pickExportDir: async () => {
      return await files.makeTempDir()
    },
    exportImages: async (op) => {
      const result = await transport.invoke('slides:export-images', op)
      if (!result || typeof result !== 'object' || (result as { ok?: unknown }).ok !== true) {
        return result
      }
      const paths = (result as { paths?: string[] }).paths ?? []
      for (const p of paths) {
        const file = await files.readFileBytes(p)
        downloadBytes(file.name, file.bytes)
      }
      return result
    },
    pickExportPdfPath: async (defaultName) => {
      const dir = await files.makeTempDir()
      const safeName = String(defaultName || 'export.pdf').replace(/[/\\:*?"<>|]/g, '_')
      return `${dir}/${safeName}`
    },
    exportPdf: async (op) => {
      const result = await transport.invoke('slides:export-pdf', op)
      if (!result || typeof result !== 'object' || (result as { ok?: unknown }).ok !== true) {
        return result
      }
      const path = (result as { path?: string }).path
      if (path) {
        const file = await files.readFileBytes(path)
        downloadBytes(file.name, file.bytes)
      }
      return result
    },
  })
  bridgedWindow.desktop = createSlidesFilesApi(transport, {})
  bridgedWindow.projectApi = createSlidesProjectApi(transport)
}

async function deckSizePx(): Promise<{ cx: number; cy: number } | null> {
  const size = await (window as unknown as { slidesApi?: { getSlideSize?: () => Promise<unknown> } })
    .slidesApi?.getSlideSize?.()
  if (size && typeof size === 'object') {
    const record = size as { cx?: unknown; cy?: unknown }
    if (typeof record.cx === 'number' && typeof record.cy === 'number') {
      return { cx: record.cx, cy: record.cy }
    }
  }
  return null
}

function imageNaturalSize(bytes: ArrayBuffer): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(new Blob([bytes]))
    const img = new Image()
    img.onload = () => {
      const size = { width: img.naturalWidth || 4, height: img.naturalHeight || 3 }
      URL.revokeObjectURL(url)
      resolve(size)
    }
    img.onerror = () => {
      URL.revokeObjectURL(url)
      resolve({ width: 4, height: 3 })
    }
    img.src = url
  })
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
