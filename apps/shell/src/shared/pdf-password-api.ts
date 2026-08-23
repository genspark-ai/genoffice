// Contract between the PDF password prompt window (pdf-password.html) and the
// shell main process (P23). IPC surface: get-state / submit / cancel + a
// state-changed push event, mirroring update-api.ts.

export const PDF_PASSWORD_CHANNELS = {
  getState: 'pdf-password:get-state',
  submit: 'pdf-password:submit',
  cancel: 'pdf-password:cancel',
  changed: 'pdf-password:changed',
} as const

/** window copy is localized in the main process (owner of UI language) */
export interface PdfPasswordUiStrings {
  title: string
  /** first attempt: "this PDF is encrypted, enter the password" */
  prompt: string
  /** retry attempts: "wrong password, try again" */
  retryPrompt: string
  ok: string
  cancel: string
  /** shown while a submitted password is being verified by the converter */
  verifying: string
}

export interface PdfPasswordUiState {
  /** base name of the PDF being converted, shown under the title */
  fileName: string
  /** true after a submitted password was rejected → show retryPrompt */
  retry: boolean
  /** true while the main process is re-running the conversion */
  busy: boolean
  /** BCP-47 tag for documentElement.lang (drives CJK font selection) */
  lang: string
  strings: PdfPasswordUiStrings
}

export interface PdfPasswordWindowApi {
  getState(): Promise<PdfPasswordUiState | null>
  submit(password: string): void
  cancel(): void
  onState(handler: (state: PdfPasswordUiState) => void): () => void
}
