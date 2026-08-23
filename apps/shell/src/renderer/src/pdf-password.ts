import type { PdfPasswordUiState, PdfPasswordWindowApi } from '../../shared/pdf-password-api'

// exposed by src/preload/pdf-password.ts
const api = (window as unknown as { aiOfficePdfPassword: PdfPasswordWindowApi }).aiOfficePdfPassword

const el = (id: string): HTMLElement => document.getElementById(id) as HTMLElement
const title = el('title')
const file = el('file')
const prompt = el('prompt')
const label = el('label')
const pwdWrap = el('pwd-wrap')
const password = el('password') as HTMLInputElement
const eye = el('eye') as HTMLButtonElement
const error = el('error')
const errorText = el('error-text')
const ok = el('ok') as HTMLButtonElement
const cancel = el('cancel') as HTMLButtonElement

let busy = false
let showLabel = ''
let hideLabel = ''

function render(state: PdfPasswordUiState): void {
  busy = state.busy
  const s = state.strings
  showLabel = s.show
  hideLabel = s.hide

  document.documentElement.lang = state.lang
  document.title = s.title
  title.textContent = s.title
  file.textContent = state.fileName
  file.title = state.fileName
  prompt.textContent = state.busy ? s.verifying : s.prompt
  label.textContent = s.label
  password.placeholder = s.placeholder
  ok.textContent = s.ok
  cancel.textContent = s.cancel

  const failed = state.retry && !state.busy
  // a rejected attempt comes back with retry=true — clear it for the next try
  if (failed) password.value = ''
  password.classList.toggle('invalid', failed)
  pwdWrap.classList.toggle('has-error', failed)
  error.classList.toggle('visible', failed)
  errorText.textContent = failed ? s.retryPrompt : ''

  password.disabled = state.busy
  eye.disabled = state.busy
  eye.setAttribute('aria-label', eye.classList.contains('showing') ? hideLabel : showLabel)
  ok.disabled = state.busy || !password.value
  cancel.disabled = state.busy
  if (!state.busy) password.focus()
}

function submit(): void {
  if (busy || !password.value) return
  api.submit(password.value)
}

ok.addEventListener('click', submit)
cancel.addEventListener('click', () => {
  if (!busy) api.cancel()
})
password.addEventListener('input', () => {
  ok.disabled = busy || !password.value
  // typing clears the previous failure styling
  password.classList.remove('invalid')
  pwdWrap.classList.remove('has-error')
  error.classList.remove('visible')
})
password.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') submit()
})
// keep the caret in the password field: without this, clicking the toggle
// focuses the button and further typing goes nowhere
eye.addEventListener('mousedown', (e) => e.preventDefault())
eye.addEventListener('click', () => {
  const show = password.type === 'password'
  const caret: [number, number] = [
    password.selectionStart ?? password.value.length,
    password.selectionEnd ?? password.value.length,
  ]
  password.type = show ? 'text' : 'password'
  eye.classList.toggle('showing', show)
  eye.setAttribute('aria-label', show ? hideLabel : showLabel)
  // the type swap resets the caret to 0 — restore it now and again on the
  // next frame, since Chromium clears it once more asynchronously
  const restore = (): void => {
    if (document.activeElement === password) password.setSelectionRange(caret[0], caret[1])
  }
  restore()
  requestAnimationFrame(restore)
})
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !busy) api.cancel()
})

api.onState(render)
void api.getState().then((state) => {
  if (state) render(state)
})
