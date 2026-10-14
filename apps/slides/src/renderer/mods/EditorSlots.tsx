import { useEffect, useMemo, useSyncExternalStore } from 'react'
import type { FeatureRecord, SlotPlace } from '@genoffice/agent-core'
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

/** Re-read the registry; concurrent calls share one round trip. */
export function refreshEditorFeatures(): Promise<void> {
  if (inflight) return inflight
  const bridge = typeof window === 'undefined' ? undefined : window.slidesApi
  if (typeof bridge?.featuresList !== 'function') {
    cache = []
    emit()
    return Promise.resolve()
  }
  const run = bridge
    .featuresList()
    .then((list) => {
      cache = list
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
}: {
  place: SlotPlace
  onAsk(prompt: string): void
}): React.JSX.Element | null {
  const features = useEditorFeatures()
  useEffect(() => {
    void refreshEditorFeatures()
  }, [])
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
