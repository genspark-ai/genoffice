// Contract between the update window renderer (update.html) and the shell
// main process. Update-dialog IPC surface:
// get-state / download / install / later + a state-changed push event.

export const UPDATE_CHANNELS = {
  getState: 'update:get-state',
  download: 'update:download',
  install: 'update:install',
  later: 'update:later',
  openDownload: 'update:open-download',
  changed: 'update:changed',
} as const

/** 'manual' = automatic updating is not working for this version (repeated
 * failures, e.g. a signing-identity change the installed app refuses) — the
 * dialog guides the user to download the installer from the releases page */
export type UpdatePhase = 'available' | 'downloading' | 'downloaded' | 'error' | 'manual'

/** window copy is localized in the main process (owner of UI language) */
export interface UpdateUiStrings {
  title: string
  headline: string
  desc: string
  download: string
  later: string
  install: string
  downloading: string
  failed: string
  retry: string
  manualDesc: string
  openDownload: string
}

export interface UpdateUiState {
  phase: UpdatePhase
  version: string
  currentVersion: string
  /** 0-100, meaningful while downloading */
  percent: number
  /** BCP-47 tag for documentElement.lang (drives CJK font selection) */
  lang: string
  strings: UpdateUiStrings
}

export interface UpdateWindowApi {
  getState(): Promise<UpdateUiState | null>
  download(): void
  install(): void
  later(): void
  openDownload(): void
  onState(handler: (state: UpdateUiState) => void): () => void
}

/** user-selectable update channel; 'stable' maps to the latest.yml feed, 'beta' to beta.yml */
export type UpdateChannel = 'stable' | 'beta'

export function isUpdateChannel(v: unknown): v is UpdateChannel {
  return v === 'stable' || v === 'beta'
}
