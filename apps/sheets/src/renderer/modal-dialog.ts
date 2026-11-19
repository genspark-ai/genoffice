import { useEffect, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react'

const FOCUSABLE =
  'button:not([disabled]), input:not([disabled]), textarea:not([disabled]), select:not([disabled]), a[href], [tabindex]:not([tabindex="-1"])'

export interface ModalDialogProps {
  readonly ref: RefObject<HTMLDivElement | null>
  readonly 'aria-modal': true
  readonly onKeyDown: (event: ReactKeyboardEvent<HTMLDivElement>) => void
}

/// Excel-style modal contract for the Sheets dialogs: Escape closes, Tab
/// cycles inside the card, the first control takes focus on open and the
/// opener gets it back on close. Spread the result onto the role="dialog" node.
export function useModalDialog(onClose: () => void): ModalDialogProps {
  const ref = useRef<HTMLDivElement>(null)
  const closeRef = useRef(onClose)
  closeRef.current = onClose
  // Read during render: an autoFocus child already owns focus once effects run
  const [opener] = useState(() => document.activeElement as HTMLElement | null)

  useEffect(() => {
    const card = ref.current
    // An autoFocus child may already own focus; leave it there
    if (card && !card.contains(document.activeElement)) {
      ;(card.querySelector<HTMLElement>(FOCUSABLE) ?? card).focus()
    }
    return () => {
      if (opener && opener !== document.body && opener.isConnected) opener.focus()
    }
  }, [opener])

  // Clicking non-focusable card content drops focus to body, where the React handler never sees Escape/Tab.
  useEffect(() => {
    const onKey = (event: KeyboardEvent): void => {
      if (ref.current?.contains(document.activeElement)) return
      if (event.key === 'Escape') {
        event.preventDefault()
        closeRef.current()
      } else if (event.key === 'Tab') {
        trapTab(ref.current, event)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      closeRef.current()
    } else if (event.key === 'Tab') {
      trapTab(ref.current, event)
    }
  }

  return { ref, 'aria-modal': true, onKeyDown }
}

export function trapTab(
  card: HTMLElement | null,
  event: { readonly shiftKey: boolean; preventDefault: () => void },
): void {
  if (!card) return
  const items = [...card.querySelectorAll<HTMLElement>(FOCUSABLE)]
  if (items.length === 0) {
    event.preventDefault()
    return
  }
  const first = items[0]!
  const last = items[items.length - 1]!
  const active = document.activeElement
  if (event.shiftKey && (active === first || !card.contains(active))) {
    event.preventDefault()
    last.focus()
  } else if (!event.shiftKey && (active === last || !card.contains(active))) {
    event.preventDefault()
    first.focus()
  }
}
