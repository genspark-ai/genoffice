import { useEffect, useRef } from 'react'

/** Same cadence as the Docs autosave (30s tick + save on window blur) */
export const AUTOSAVE_INTERVAL_MS = 30_000

/**
 * Autosave scheduling (Docs strategy): calls `save` every `intervalMs` and whenever
 * the window loses focus, but only while `shouldSave()` reports pending changes.
 * Returns a stop function that detaches the timer and the blur listener.
 */
export function startAutosave(
  shouldSave: () => boolean,
  save: () => void,
  intervalMs: number = AUTOSAVE_INTERVAL_MS,
): () => void {
  const tick = () => {
    if (shouldSave()) save()
  }
  const id = window.setInterval(tick, intervalMs)
  window.addEventListener('blur', tick)
  return () => {
    window.clearInterval(id)
    window.removeEventListener('blur', tick)
  }
}

/** React binding: installs the scheduler once and reads the latest callbacks through a ref */
export function useAutosave(shouldSave: () => boolean, save: () => void): void {
  const latest = useRef({ shouldSave, save })
  latest.current = { shouldSave, save }
  useEffect(
    () =>
      startAutosave(
        () => latest.current.shouldSave(),
        () => latest.current.save(),
      ),
    [],
  )
}
