import { useEffect, useRef } from 'react'
import type { RefObject } from 'react'

const FOCUSABLE =
  'button:not([disabled]), [href], input:not([disabled]), [tabindex]:not([tabindex="-1"])'

/**
 * Keyboard contract for a modal layered over another dialog (e.g. Settings):
 * `initial` takes focus on open and the opener gets it back on close, Tab and
 * Shift+Tab wrap inside `card`, and Escape calls `onClose` without reaching the
 * dialog underneath (handled in the capture phase, before its own listener).
 */
export function useModalFocus(
  card: RefObject<HTMLElement | null>,
  initial: RefObject<HTMLElement | null>,
  onClose: () => void,
): void {
  // the latest onClose without re-running the open/close effect on every render
  const close = useRef(onClose)
  close.current = onClose

  useEffect(() => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    initial.current?.focus()
    return () => opener?.focus()
  }, [initial])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        e.preventDefault()
        close.current()
        return
      }
      if (e.key !== 'Tab' || !card.current) return
      const items = [...card.current.querySelectorAll<HTMLElement>(FOCUSABLE)]
      if (items.length === 0) return
      const first = items[0]!
      const last = items[items.length - 1]!
      const inside = card.current.contains(document.activeElement)
      if (e.shiftKey && (!inside || document.activeElement === first)) {
        e.preventDefault()
        last.focus()
      } else if (!e.shiftKey && (!inside || document.activeElement === last)) {
        e.preventDefault()
        first.focus()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [card])
}
