import type { LocalVlmSettings, LocalVlmStatus } from '@genoffice/ai-provider/local-vlm-meta'

export type {
  LocalVlmDownloadProgress,
  LocalVlmSettings,
  LocalVlmStatus,
} from '@genoffice/ai-provider/local-vlm-meta'

export const LOCAL_VLM_CHANNELS = {
  getStatus: 'local-vlm:get-status',
  setSettings: 'local-vlm:set-settings',
  download: 'local-vlm:download',
  cancelDownload: 'local-vlm:cancel-download',
  start: 'local-vlm:start',
  stop: 'local-vlm:stop',
  test: 'local-vlm:test',
  pickFile: 'local-vlm:pick-file',
  statusChanged: 'local-vlm:status-changed',
} as const

export type LocalVlmPickKind = 'model' | 'mmproj' | 'runtime'

export interface LocalVlmApi {
  getLocalVlmStatus(): Promise<LocalVlmStatus | null>
  setLocalVlmSettings(patch: Partial<LocalVlmSettings>): Promise<LocalVlmStatus | null>
  downloadLocalVlm(options?: {
    includeRuntime?: boolean
    includeModel?: boolean
  }): Promise<{ ok: boolean; status?: LocalVlmStatus; error?: string }>
  cancelLocalVlmDownload(): Promise<boolean>
  startLocalVlm(): Promise<{ ok: boolean; status?: LocalVlmStatus; error?: string }>
  stopLocalVlm(): Promise<LocalVlmStatus | null>
  testLocalVlm(): Promise<{ ok: boolean; error?: string; latencyMs?: number }>
  pickLocalVlmFile(kind: LocalVlmPickKind): Promise<string | null>
  onLocalVlmStatusChanged(handler: (status: LocalVlmStatus) => void): () => void
}
