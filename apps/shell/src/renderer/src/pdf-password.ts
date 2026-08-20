import type { PdfPasswordUiState, PdfPasswordWindowApi } from '../../shared/pdf-password-api'

// exposed by src/preload/pdf-password.ts
const api = (window as unknown as { aiOfficePdfPassword: PdfPasswordWindowApi }).aiOfficePdfPassword

const el = (id: string): HTMLElement => document.getElementById(id) as HTMLElement
const title = el('title')
const file = el('file')
const prompt = el('prompt')
const password = el('password') as HTMLInputElement
const ok = el('ok') as HTMLButtonElement
const cancel = el('cancel') as HTMLButtonElement

let busy = false

function render(state: PdfPasswordUiState): void {
  busy = state.busy
  const s = state.strings

  document.documentElement.lang = state.lang
  document.title = s.title
  title.textContent = s.title
  file.textContent = state.fileName
  file.title = state.fileName
  prompt.textContent = state.busy ? s.verifying : state.retry ? s.retryPrompt : s.prompt
  prompt.classList.toggle('error', state.retry && !state.busy)
  ok.textContent = s.ok
  cancel.textContent = s.cancel

  password.disabled = state.busy
  ok.disabled = state.busy
  cancel.disabled = state.busy
  if (!state.busy) {
    // a rejected attempt comes back with retry=true — clear it for the next try
    if (state.retry) password.value = ''
    password.focus()
  }
}

function submit(): void {
  if (busy) return
  api.submit(password.value)
}

ok.addEventListener('click', submit)
cancel.addEventListener('click', () => {
  if (!busy) api.cancel()
})
password.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') submit()
})
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !busy) api.cancel()
})

api.onState(render)
void api.getState().then((state) => {
  if (state) render(state)
})
