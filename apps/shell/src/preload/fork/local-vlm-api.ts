import type { IpcRenderer, IpcRendererEvent } from 'electron'
import {
  LOCAL_VLM_CHANNELS,
  type LocalVlmApi,
  type LocalVlmPickKind,
  type LocalVlmSettings,
  type LocalVlmStatus,
} from '../../shared/fork/local-vlm'

export function createLocalVlmPreloadApi(ipcRenderer: IpcRenderer): LocalVlmApi {
  return {
    async getLocalVlmStatus() {
      return (
        ((await ipcRenderer.invoke(LOCAL_VLM_CHANNELS.getStatus)) as LocalVlmStatus | null) ?? null
      )
    },
    async setLocalVlmSettings(patch: Partial<LocalVlmSettings>) {
      if (!patch || typeof patch !== 'object') throw new Error('Invalid settings.')
      return (
        ((await ipcRenderer.invoke(
          LOCAL_VLM_CHANNELS.setSettings,
          patch,
        )) as LocalVlmStatus | null) ?? null
      )
    },
    async downloadLocalVlm(options) {
      return (await ipcRenderer.invoke(LOCAL_VLM_CHANNELS.download, options ?? {})) as {
        ok: boolean
        status?: LocalVlmStatus
        error?: string
      }
    },
    async cancelLocalVlmDownload() {
      return (await ipcRenderer.invoke(LOCAL_VLM_CHANNELS.cancelDownload)) === true
    },
    async startLocalVlm() {
      return (await ipcRenderer.invoke(LOCAL_VLM_CHANNELS.start)) as {
        ok: boolean
        status?: LocalVlmStatus
        error?: string
      }
    },
    async stopLocalVlm() {
      return ((await ipcRenderer.invoke(LOCAL_VLM_CHANNELS.stop)) as LocalVlmStatus | null) ?? null
    },
    async testLocalVlm() {
      return (await ipcRenderer.invoke(LOCAL_VLM_CHANNELS.test)) as {
        ok: boolean
        error?: string
        latencyMs?: number
      }
    },
    async pickLocalVlmFile(kind: LocalVlmPickKind) {
      const result: unknown = await ipcRenderer.invoke(LOCAL_VLM_CHANNELS.pickFile, kind)
      return typeof result === 'string' && result ? result : null
    },
    onLocalVlmStatusChanged(handler) {
      const listener = (_event: IpcRendererEvent, status: unknown) => {
        if (status && typeof status === 'object') handler(status as LocalVlmStatus)
      }
      ipcRenderer.on(LOCAL_VLM_CHANNELS.statusChanged, listener)
      return () => ipcRenderer.removeListener(LOCAL_VLM_CHANNELS.statusChanged, listener)
    },
  }
}
