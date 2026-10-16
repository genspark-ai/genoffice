import { useSyncExternalStore } from 'react'
import { activeAppearance, type FeatureRecord } from '@genoffice/agent-core'

/**
 * The renderer's one copy of the feature registry.
 *
 * Home renders slots from it and the mod assistant writes through it, so a
 * feature the assistant records shows up in the interface without either side
 * knowing about the other. It reads the same IPC bridge the panel uses.
 */

let cache: FeatureRecord[] = []
let inflight: Promise<void> | null = null
const listeners = new Set<() => void>()

function emit(): void {
  for (const listener of listeners) listener()
}

export function getFeatures(): FeatureRecord[] {
  return cache
}

/**
 * The chrome a feature asked for (K8). The registry's newest appearance wins;
 * removing it hands the shell back its own defaults. This only re-points chrome
 * variables — a document surface, a page background and the print/export path
 * never see it.
 */
function applyAppearance(features: FeatureRecord[]): void {
  if (typeof document === 'undefined') return
  const chosen = activeAppearance(features)
  const root = document.documentElement
  const set = (name: string, value: string | undefined): void => {
    if (value && value !== 'default') root.setAttribute(name, value)
    else root.removeAttribute(name)
  }
  set('data-slot-background', chosen?.background)
  set('data-slot-accent', chosen?.accent)
}

export function subscribeFeatures(listener: () => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}

/** Re-read the registry; concurrent calls share one round trip. */
export function refreshFeatures(): Promise<void> {
  if (inflight) return inflight
  const bridge = typeof window === 'undefined' ? undefined : window.aiOffice
  if (typeof bridge?.featuresList !== 'function') {
    cache = []
    emit()
    return Promise.resolve()
  }
  inflight = bridge
    .featuresList()
    .then((list) => {
      cache = list
      applyAppearance(list)
      emit()
    })
    .catch(() => {
      cache = []
      applyAppearance([])
      emit()
    })
    .finally(() => {
      inflight = null
    })
  return inflight
}

export function useFeatures(): FeatureRecord[] {
  return useSyncExternalStore(subscribeFeatures, getFeatures)
}

export function featureById(id: string): FeatureRecord | null {
  return cache.find((feature) => feature.id === id) ?? null
}
