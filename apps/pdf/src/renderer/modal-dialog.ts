import { useEffect, useRef, useState } from 'react'
import type { RefObject } from 'react'

/**
 * Modal contract shared by the .pdf-modal dialogs: Escape cancels (unless the
 * dialog hands the key to a nested popover), the first field takes focus on
 * open and the opener gets it back on close. Put the ref on the dialog card.
 */
export function useModalDialog(
  onCancel: () => void,
  { escape = true }: { escape?: boolean } = {},
): RefObject<HTMLDivElement | null> {
  const dialogRef = useRef<HTMLDivElement>(null)
  // Read during render: an autoFocus child already owns focus once effects run
  const [opener] = useState(() => document.activeElement as HTMLElement | null)

  useEffect(() => {
    if (!escape) return
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        event.preventDefault()
        onCancel()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onCancel, escape])

  useEffect(() => {
    const root = dialogRef.current
    if (root && !root.contains(document.activeElement)) {
      // Tab strips and thumbnails precede the text field in some dialogs; land in the field so typing works
      ;(
        root.querySelector<HTMLElement>('input, textarea, select') ??
        root.querySelector<HTMLElement>('button')
      )?.focus()
    }
    return () => {
      if (opener && opener !== document.body && opener.isConnected) opener.focus()
    }
  }, [opener])

  return dialogRef
}
