// Contract between the update window renderer (update.html) and the shell
// main process. Update-dialog IPC surface:
// get-state / download / install / later + a state-changed push event.

export const UPDATE_CHANNELS = {
  getState: 'update:get-state',
  download: 'update:download',
  install: 'update:install',
  later: 'update:later',
  changed: 'update:changed',
} as const

export type UpdatePhase = 'available' | 'downloading' | 'downloaded' | 'error'

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
  onState(handler: (state: UpdateUiState) => void): () => void
}
