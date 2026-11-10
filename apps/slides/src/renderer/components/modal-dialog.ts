import { useEffect, useId, useRef, useState } from 'react'
import type { RefObject } from 'react'
import { useEscOverlay } from '../esc-overlay'

export interface ModalDialogProps {
  ref: RefObject<HTMLDivElement | null>
  role: 'dialog'
  'aria-modal': true
  'aria-labelledby': string
}

/**
 * Dialog contract shared by the .modal cards (same behaviour as PrintDialog):
 * Escape closes through the overlay registry, the first control takes focus
 * on open, the opener gets it back on close, and the card is announced by its
 * heading. Spread `dialogProps` onto the card and give the <h2> `titleId`.
 */
export function useModalDialog(onClose: () => void): {
  titleId: string
  dialogProps: ModalDialogProps
} {
  useEscOverlay(true, onClose)
  const ref = useRef<HTMLDivElement>(null)
  const titleId = useId()
  // Read during render: an autoFocus child already owns focus once effects run
  const [opener] = useState(() => document.activeElement as HTMLElement | null)

  useEffect(() => {
    const card = ref.current
    if (card && !card.contains(document.activeElement)) {
      card.querySelector<HTMLElement>('input, textarea, select, button')?.focus()
    }
    return () => {
      if (opener && opener !== document.body && opener.isConnected) opener.focus()
    }
  }, [opener])

  return {
    titleId,
    dialogProps: { ref, role: 'dialog', 'aria-modal': true, 'aria-labelledby': titleId },
  }
}
