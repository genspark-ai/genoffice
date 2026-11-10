/**
 * Shared modal keyboard behavior for the renderer dialogs: Esc closes (stopped
 * so it never reaches global listeners), Tab cycles inside the modal, and the
 * first form control gets focus on mount unless something inside is already
 * focused. Spread the returned ref/onKeyDown onto the modal backdrop element.
 */
import { useEffect, useRef } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'

const FOCUSABLE = 'a[href], button, input, textarea, select, [tabindex]:not([tabindex="-1"])'

export function useModalKeys(onClose: () => void, options?: { restoreFocus?: boolean }) {
  const ref = useRef<HTMLDivElement>(null)
  const restoreFocus = options?.restoreFocus ?? false

  useEffect(() => {
    const el = ref.current
    if (!el) return
    const previous =
      restoreFocus && document.activeElement instanceof HTMLElement ? document.activeElement : null
    if (!el.contains(document.activeElement)) {
      // skip disabled controls, the same way trapTab does: a modal whose only
      // early control is disabled would otherwise never take focus at all
      const first = [...el.querySelectorAll<HTMLElement>('input, textarea, select, button')].find(
        (c) => !c.hasAttribute('disabled'),
      )
      ;(first ?? el).focus()
    }
    return () => {
      if (previous && previous.isConnected) previous.focus()
    }
  }, [restoreFocus])

  const onKeyDown = (e: ReactKeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault()
      e.stopPropagation()
      onClose()
      return
    }
    if (e.key === 'Tab') trapTab(ref.current, e)
  }

  return { ref, onKeyDown }
}

/** Cycle Tab/Shift+Tab inside a container (pure DOM helper, exported for tests). */
export function trapTab(
  container: HTMLElement | null,
  e: { key: string; shiftKey: boolean; preventDefault: () => void },
): void {
  if (!container || e.key !== 'Tab') return
  const items = [...container.querySelectorAll<HTMLElement>(FOCUSABLE)].filter(
    (el) => !el.hasAttribute('disabled') && el.tabIndex >= 0,
  )
  if (items.length === 0) {
    e.preventDefault()
    return
  }
  const first = items[0]!
  const last = items[items.length - 1]!
  const active = document.activeElement as HTMLElement | null
  // A backdrop that holds focus itself is inside the trap, so name it alongside
  // the first/last control: without it Shift+Tab from the backdrop is left to
  // the browser and walks out of the modal.
  if (e.shiftKey && (active === first || active === container || !container.contains(active))) {
    e.preventDefault()
    last.focus()
  } else if (!e.shiftKey && (active === last || active === container)) {
    e.preventDefault()
    first.focus()
  }
}
