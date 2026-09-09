import { useSyncExternalStore } from 'react'
import {
  DEFAULT_AI_PANEL_PREFS,
  aiPanelZoom,
  normalizeAiPanelPrefs,
  sameAiPanelPrefs,
  type AiPanelPrefs,
} from './ai-panel-prefs'

let current: AiPanelPrefs = DEFAULT_AI_PANEL_PREFS
const listeners = new Set<() => void>()

function subscribe(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

/**
 * Renderer-side entry for the shell's AI panel preferences: mirrors the font
 * size onto `<html data-ai-font-size>` plus `--ai-font-zoom` (see
 * ai-panel-prefs.css) and feeds `useAiPanelPrefs()` consumers such as the
 * composer's spellcheck flag.
 */
export function applyAiPanelPrefs(raw: unknown): void {
  const next = normalizeAiPanelPrefs(raw)
  if (sameAiPanelPrefs(next, current)) return
  current = next
  const html = document.documentElement
  if (next.fontSize === 'default') {
    delete html.dataset.aiFontSize
    html.style.removeProperty('--ai-font-zoom')
  } else {
    html.dataset.aiFontSize = next.fontSize
    html.style.setProperty('--ai-font-zoom', String(aiPanelZoom(next)))
  }
  for (const listener of listeners) listener()
}

export function useAiPanelPrefs(): AiPanelPrefs {
  return useSyncExternalStore(
    subscribe,
    () => current,
    () => DEFAULT_AI_PANEL_PREFS,
  )
}
