/**
 * Modal focus handling for the shared ImageViewer. Deliberately viewer-local:
 * the editors' dialogs have their own shared modal helper, and the viewer
 * should not have to wait on (or inherit the crop behaviour of) it.
 *
 * Spread the returned `ref` / `onKeyDown` / `onBlur` onto the viewer backdrop
 * (which needs `tabIndex={-1}` so it can hold focus itself): Esc closes
 * (stopped, so it never reaches the editor behind the overlay), Tab cycles
 * inside the backdrop, the control named by `initialFocus` takes focus on open,
 * focus that drops to body comes back onto the backdrop, and focus returns to
 * whatever had it when the viewer opened.
 */
import { useEffect, useRef } from 'react'
import type { FocusEvent as ReactFocusEvent, KeyboardEvent as ReactKeyboardEvent } from 'react'

const FOCUSABLE = 'button, input, textarea, select, [tabindex]:not([tabindex="-1"])'

export function useImageViewerFocus(onClose: () => void, initialFocus: string) {
  const ref = useRef<HTMLDivElement>(null)
  // The restore target is captured once, on the first commit that has a mounted
  // backdrop. An effect that re-ran would otherwise record the viewer's own
  // control as the "previous" element and restore focus to the closed dialog.
  const restoreTo = useRef<HTMLElement | null>(null)
  const captured = useRef(false)

  useEffect(() => {
    const el = ref.current
    if (!el) return
    if (!captured.current) {
      captured.current = true
      if (document.activeElement instanceof HTMLElement) restoreTo.current = document.activeElement
    }
    // an element focused during commit (autoFocus) already counts as "inside"
    if (!el.contains(document.activeElement)) {
      const target = el.querySelector<HTMLElement>(initialFocus)
      ;(target ?? el).focus()
    }
    return () => {
      const previous = restoreTo.current
      if (previous && previous.isConnected) previous.focus()
    }
  }, [initialFocus])

  // Clicking a non-focusable part of the toolbar (zoom %, separators) blurs
  // focus to body, where the backdrop's React handler never sees Esc. Catch it
  // on window (capture, so the editor behind the overlay still never does).
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape' || !focusOutsideViewer(ref.current)) return
      e.preventDefault()
      e.stopPropagation()
      onClose()
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose])

  const onKeyDown = (e: ReactKeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault()
      e.stopPropagation()
      onClose()
      return
    }
    if (e.key === 'Tab') trapTab(ref.current, e)
  }

  /** Focus dropping to body (nothing else took it) would let Tab walk into the
   * editor: pull it back onto the backdrop. */
  const onBlur = (e: ReactFocusEvent) => {
    const el = ref.current
    if (!el || e.relatedTarget) return
    el.focus()
  }

  return { ref, onKeyDown, onBlur }
}

function focusOutsideViewer(container: HTMLElement | null): boolean {
  const active = document.activeElement
  return !active || active === document.body || !container?.contains(active)
}

/** Cycle Tab / Shift+Tab inside the backdrop so focus cannot walk out. */
function trapTab(container: HTMLElement | null, e: ReactKeyboardEvent): void {
  if (!container) return
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
  if (e.shiftKey && (active === first || !container.contains(active))) {
    e.preventDefault()
    last.focus()
  } else if (!e.shiftKey && active === last) {
    e.preventDefault()
    first.focus()
  }
}
