import type { UpdateUiState, UpdateWindowApi } from '../../shared/update-api'

// exposed by src/preload/update.ts
const api = (window as unknown as { aiOfficeUpdate: UpdateWindowApi }).aiOfficeUpdate

const el = (id: string): HTMLElement => document.getElementById(id) as HTMLElement
const headline = el('headline')
const verCurrent = el('ver-current')
const verNew = el('ver-new')
const desc = el('desc')
const progress = el('progress')
const bar = el('bar')
const progressText = el('progress-text')
const percentText = el('percent')
const action = el('action') as HTMLButtonElement
const later = el('later') as HTMLButtonElement

let phase: UpdateUiState['phase'] = 'available'

function render(state: UpdateUiState): void {
  phase = state.phase
  const s = state.strings

  document.documentElement.lang = state.lang
  document.title = s.title
  headline.textContent = s.headline
  verCurrent.textContent = `v${state.currentVersion}`
  verNew.textContent = `v${state.version}`
  later.textContent = s.later

  desc.classList.toggle('error', state.phase === 'error')

  switch (state.phase) {
    case 'available':
      desc.textContent = s.desc
      progress.style.display = 'none'
      action.style.display = ''
      action.textContent = s.download
      break
    case 'downloading': {
      desc.textContent = s.desc
      progress.style.display = 'flex'
      action.style.display = 'none'
      const pct = Math.max(0, Math.min(100, Math.round(state.percent)))
      bar.style.width = `${pct}%`
      progressText.textContent = s.downloading
      percentText.textContent = `${pct}%`
      break
    }
    case 'downloaded':
      desc.textContent = s.desc
      progress.style.display = 'none'
      action.style.display = ''
      action.textContent = s.install
      break
    case 'error':
      desc.textContent = s.failed
      progress.style.display = 'none'
      action.style.display = ''
      action.textContent = s.retry
      break
  }
}

action.addEventListener('click', () => {
  if (phase === 'downloaded') api.install()
  else api.download()
})
later.addEventListener('click', () => api.later())

api.onState(render)
void api.getState().then((state) => {
  if (state) render(state)
})
