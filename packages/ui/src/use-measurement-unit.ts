/**
 * An editor window's copy of the app-wide measurement unit (Settings ▸
 * General), for editors that only follow the setting. Connect once at startup
 * with the preload's getter and change event; components read it with
 * useMeasurementUnit() and re-render when it changes in Settings.
 */
import { useSyncExternalStore } from 'react'

import { defaultMeasurementUnit, isMeasurementUnitState, type MeasurementUnit } from './measurement'

let current: MeasurementUnit | null = null
const listeners = new Set<() => void>()

/** the unit in effect: the shell's, or the locale's default when running standalone */
export function measurementUnit(): MeasurementUnit {
  return current ?? defaultMeasurementUnit(globalThis.navigator?.language ?? 'en-US')
}

function apply(state: unknown): void {
  if (!isMeasurementUnitState(state) || state.unit === current) return
  current = state.unit
  for (const l of listeners) l()
}

export interface MeasurementUnitSource {
  get(): Promise<unknown>
  onChanged(handler: (state: unknown) => void): () => void
}

/** follow the shell's unit, including later changes made in Settings */
export function connectMeasurementUnit(source: MeasurementUnitSource): () => void {
  const off = source.onChanged(apply)
  void source
    .get()
    .then(apply)
    .catch(() => undefined)
  return off
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}

export function useMeasurementUnit(): MeasurementUnit {
  return useSyncExternalStore(subscribe, measurementUnit, measurementUnit)
}

/** test hook: forget the shell's answer */
export function resetMeasurementUnitForTests(): void {
  current = null
}
