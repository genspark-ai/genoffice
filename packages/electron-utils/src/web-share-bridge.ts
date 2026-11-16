/// Preload-side bridge for the status-bar web-share chips (every editor).
/// The channels are registered once by the shell (the web service is
/// process-wide); this exposes the three calls plus the change push as one
/// `window.webShareApi` surface. Install alongside installDropOpenBridge.
import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'

export const WEB_SHARE_CHANNELS = {
  status: 'web-share:status',
  start: 'web-share:start',
  stop: 'web-share:stop',
  changed: 'web-share:changed',
} as const

export interface WebShareBridgeStatus {
  running: boolean
  url?: string
  localUrl?: string
  port?: number
  error?: string
}

export function installWebShareBridge(): void {
  contextBridge.exposeInMainWorld('webShareApi', {
    status: (): Promise<WebShareBridgeStatus> => ipcRenderer.invoke(WEB_SHARE_CHANNELS.status),
    start: (): Promise<WebShareBridgeStatus> => ipcRenderer.invoke(WEB_SHARE_CHANNELS.start),
    stop: (): Promise<WebShareBridgeStatus> => ipcRenderer.invoke(WEB_SHARE_CHANNELS.stop),
    onChanged: (handler: (status: WebShareBridgeStatus) => void): (() => void) => {
      const listener = (_event: IpcRendererEvent, status: WebShareBridgeStatus): void =>
        handler(status)
      ipcRenderer.on(WEB_SHARE_CHANNELS.changed, listener)
      return () => ipcRenderer.removeListener(WEB_SHARE_CHANNELS.changed, listener)
    },
  })
}
