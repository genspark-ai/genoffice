import { useEffect, useRef } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'

/**
 * Shared modal keyboard behavior: Esc closes (stopped so it never reaches
 * global listeners like "exit read mode"), and the first form control gets
 * focus on mount unless something inside is already focused (autoFocus).
 * Spread the returned ref/onKeyDown onto the modal backdrop element.
 */
export function useModalKeys(onClose: () => void) {
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const el = ref.current
    if (!el || el.contains(document.activeElement)) return
    const first = el.querySelector<HTMLElement>('input, textarea, select, button')
    ;(first ?? el).focus()
  }, [])

  const onKeyDown = (e: ReactKeyboardEvent) => {
    if (e.key !== 'Escape') return
    e.preventDefault()
    e.stopPropagation()
    onClose()
  }

  return { ref, onKeyDown }
}
