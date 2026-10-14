import { useEffect, useMemo, useSyncExternalStore } from 'react'
import {
  activeAppearance,
  markdownCustomThemeVars,
  type FeatureRecord,
  type SlotPlace,
} from '@genoffice/agent-core'
import { resolvePaperTheme, type MarkdownTheme } from '../../shared/markdown-theme'
import { SlotHost, type SlotActions } from '@genoffice/ui'

/**
 * Editor slots: the same renderer the shell uses, bound to this editor's bridge.
 *
 * The registry lives in the shell process, so the editor reads it with one call
 * and caches it; a feature the mod assistant records shows up here the next time
 * the editor refreshes. What belongs to the editor is what a press may do — hand
 * a prompt to its own assistant. This editor has no file-opener on its bridge, so
 * an open action draws disabled rather than doing nothing.
 */

let cache: FeatureRecord[] = []
let inflight: Promise<void> | null = null
const listeners = new Set<() => void>()

function emit(): void {
  for (const listener of listeners) listener()
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}

function current(): FeatureRecord[] {
  return cache
}

/** the custom-theme properties this module owns on the root element */
const CUSTOM_THEME_PROPS = [
  '--md-custom-paper',
  '--md-custom-ink',
  '--md-custom-accent',
  '--md-custom-scheme',
]

/**
 * One writer for the paper-theme attribute: the user's ribbon pick wins; only
 * when it is `default` does the mod appearance decide. A palette name re-scopes
 * the attribute; a feature's custom object additionally sets the paper/ink
 * variables the stylesheet mixes its tones from. Display only — the text,
 * saving and the print/export path never see any of it.
 */
export function applyMarkdownTheme(userTheme: MarkdownTheme): void {
  if (typeof document === 'undefined') return
  const mod = activeAppearance(current())?.markdown
  const resolved = resolvePaperTheme(userTheme, mod)
  const root = document.documentElement
  for (const name of CUSTOM_THEME_PROPS) root.style.removeProperty(name)
  root.removeAttribute('data-md-font')
  if (resolved === 'default') root.removeAttribute('data-markdown-theme')
  else root.setAttribute('data-markdown-theme', resolved)
  if (resolved === 'custom' && mod && typeof mod === 'object') {
    const { scheme, vars } = markdownCustomThemeVars(mod)
    root.style.setProperty('--md-custom-scheme', scheme)
    for (const [name, value] of Object.entries(vars)) root.style.setProperty(name, value)
    if (mod.font) root.setAttribute('data-md-font', mod.font)
  }
}

/** Re-read the registry; concurrent calls share one round trip. */
export function refreshEditorFeatures(): Promise<void> {
  if (inflight) return inflight
  const bridge = typeof window === 'undefined' ? undefined : window.markdownApi
  if (typeof bridge?.featuresList !== 'function') {
    cache = []
    emit()
    return Promise.resolve()
  }
  const run = bridge
    .featuresList()
    .then((list) => {
      cache = list
      // the editor effect below re-applies the paper theme on this change
      emit()
    })
    .catch(() => {
      cache = []
      emit()
    })
    .finally(() => {
      inflight = null
    })
  inflight = run
  return run
}

export function useEditorFeatures(): FeatureRecord[] {
  return useSyncExternalStore(subscribe, current)
}

/** The registry as of this turn, for the assistant's per-turn context. */
export function currentEditorFeatures(): FeatureRecord[] {
  return cache
}

/** The slots registered for one place in this editor. */
export function EditorSlots({
  place,
  onAsk,
  mdTheme,
}: {
  place: SlotPlace
  onAsk(prompt: string): void
  /** the user's ribbon pick; wins over what a feature asks for */
  mdTheme: MarkdownTheme
}): React.JSX.Element | null {
  const features = useEditorFeatures()
  useEffect(() => {
    void refreshEditorFeatures()
  }, [])
  useEffect(() => {
    applyMarkdownTheme(mdTheme)
  }, [mdTheme, features])
  const actions = useMemo<SlotActions>(
    () => ({
      ask: onAsk,
      // an editor has no feature panel of its own: the assistant is where a
      // feature's own view would have to live anyway
      panel: (id: string): void => onAsk('Show the feature "' + id + '".'),
    }),
    [onAsk],
  )
  return <SlotHost place={place} features={features} actions={actions} />
}
