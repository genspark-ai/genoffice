/// Browser-safe web equivalents for desktop-only channels.
///
/// The web version cannot show native dialogs, print via webContents, or touch
/// the OS clipboard/fonts directly. These helpers give the per-app web bridges
/// the same user-visible functionality with browser APIs:
///
///   • file open/pick  → `<input type="file">` + `web:write-temp-file` (bytes
///                       land in the main process, which then runs the normal
///                       open/add flow on the temp path);
///   • save/export     → main returns bytes, the browser downloads them;
///   • print           → `window.print()`;
///   • clipboard       → `navigator.clipboard`;
///   • font metrics    → canvas `measureText`;
///   • fullscreen      → `requestFullscreen`;
///   • tabs/windows    → `window.open`.
///
/// This module must stay free of node/electron imports (browser + preload-safe).

import type { IpcTransport } from './client'

/** Generic main-process web-file channels (registered by the bridge server). */
export const WEB_FILE_CHANNELS = {
  writeTempFile: 'web:write-temp-file',
  readFileBytes: 'web:read-file-bytes',
  makeTempDir: 'web:make-temp-dir',
} as const

export interface WebTempFile {
  name: string
  bytes: ArrayBuffer
}

export interface WebFileBridge {
  /** Write bytes to a main-process temp file; resolves with its path. */
  writeTempFile(name: string, bytes: ArrayBuffer): Promise<string>
  /** Read a main-process temp file back as bytes (for browser downloads). */
  readFileBytes(path: string): Promise<WebTempFile>
  /** Create a main-process temp directory; resolves with its path. */
  makeTempDir(): Promise<string>
}

export function createWebFileBridge(t: IpcTransport): WebFileBridge {
  return {
    async writeTempFile(name, bytes) {
      const result: unknown = await t.invoke(WEB_FILE_CHANNELS.writeTempFile, { name, bytes })
      if (typeof result !== 'string' || !result) {
        throw new Error('web: temp file write failed')
      }
      return result
    },
    async readFileBytes(path) {
      const result: unknown = await t.invoke(WEB_FILE_CHANNELS.readFileBytes, path)
      if (!result || typeof result !== 'object') {
        throw new Error('web: temp file read failed')
      }
      const record = result as { name?: unknown; bytes?: unknown }
      if (typeof record.name !== 'string' || !(record.bytes instanceof ArrayBuffer)) {
        throw new Error('web: temp file read returned an invalid shape')
      }
      return { name: record.name, bytes: record.bytes }
    },
    async makeTempDir() {
      const result: unknown = await t.invoke(WEB_FILE_CHANNELS.makeTempDir)
      if (typeof result !== 'string' || !result) {
        throw new Error('web: temp dir creation failed')
      }
      return result
    },
  }
}

/** Pick one or more files through a hidden `<input type="file">`. */
export function pickFileBytes(
  accept?: string,
  multiple = false,
): Promise<WebTempFile[] | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    if (accept) input.accept = accept
    input.multiple = multiple
    input.style.display = 'none'
    document.body.appendChild(input)
    input.addEventListener(
      'change',
      () => {
        const files = input.files ? Array.from(input.files) : []
        input.remove()
        if (files.length === 0) {
          resolve(null)
          return
        }
        void Promise.all(
          files.map(
            (file) =>
              new Promise<WebTempFile>((done, fail) => {
                const reader = new FileReader()
                reader.onload = () => {
                  const result = reader.result
                  const bytes =
                    result instanceof ArrayBuffer
                      ? result
                      : new Uint8Array(result as unknown as ArrayBuffer).buffer
                  done({
                    name: file.name,
                    bytes,
                  })
                }
                reader.onerror = () => fail(reader.error ?? new Error('web: file read failed'))
                reader.readAsArrayBuffer(file)
              }),
          ),
        ).then(resolve, () => resolve(null))
      },
      { once: true },
    )
    input.click()
  })
}

/** Trigger a browser download of `bytes` under `name`. */
export function downloadBytes(name: string, bytes: ArrayBuffer): void {
  const blob = new Blob([bytes])
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = name
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  setTimeout(() => URL.revokeObjectURL(url), 10_000)
}

/** Browser print dialog — the web equivalent of the desktop print pipeline. */
export function webPrint(): void {
  window.print()
}

/** Copy an image (data URL) to the clipboard via the async Clipboard API. */
export async function webCopyImage(dataUrl: string): Promise<boolean> {
  try {
    const response = await fetch(dataUrl)
    const blob = await response.blob()
    await navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })])
    return true
  } catch {
    return false
  }
}

/** Canvas-based font vertical metrics — the web equivalent of `docs:font-metrics`. */
export interface WebFontMetrics {
  ascent: number
  descent: number
  lineGap: number
  unitsPerEm: number
}

export function webFontMetrics(family: string): WebFontMetrics | null {
  if (typeof document === 'undefined') return null
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  if (!context) return null
  context.font = `64px "${family}"`
  const text = 'Hg'
  const width = context.measureText(text).width
  if (width <= 0) return null
  // 64px em box: ascent/descent approximated from the measured text extents.
  const metrics = context.measureText(text)
  const actualBoundingBoxAscent = metrics.actualBoundingBoxAscent ?? 0
  const actualBoundingBoxDescent = metrics.actualBoundingBoxDescent ?? 0
  const unitsPerEm = 1000
  const scale = unitsPerEm / 64
  return {
    ascent: Math.round(actualBoundingBoxAscent * scale),
    descent: Math.round(actualBoundingBoxDescent * scale),
    lineGap: 0,
    unitsPerEm,
  }
}

/** Browser fullscreen — the web equivalent of `slides:show-fullscreen`. */
export async function webFullscreen(): Promise<void> {
  await document.documentElement.requestFullscreen()
}

/** Open a browser tab — the web equivalent of `win:new` / `tabs:*`. */
export function webOpenTab(url: string): void {
  window.open(url, '_blank')
}
