import { useEffect, useRef, useSyncExternalStore } from 'react'

// Overlays that consume Escape from their own window listeners register here so
// the global shortcut handler (attached first, so it cannot see their
// preventDefault) leaves the key to them.
let openCount = 0
const listeners = new Set<() => void>()
const notify = (): void => listeners.forEach((l) => l())

export function useEscOverlay(open: boolean, onEscape?: () => void): void {
  const cb = useRef(onEscape)
  useEffect(() => {
    cb.current = onEscape
  }, [onEscape])
  useEffect(() => {
    if (!open) return
    openCount++
    notify()
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') cb.current?.()
    }
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('keydown', onKey)
      openCount--
      notify()
    }
  }, [open])
}

export function useEscOverlayOpen(): boolean {
  return useSyncExternalStore(
    (l) => {
      listeners.add(l)
      return () => listeners.delete(l)
    },
    () => openCount > 0,
  )
}
