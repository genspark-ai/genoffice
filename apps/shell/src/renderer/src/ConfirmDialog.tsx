import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { createPortal } from 'react-dom'
import './dialogs.css'

export interface ConfirmRequest {
  title: string
  body: string
  confirmLabel: string
  /** a destructive action (sign out, reset) gets the red confirm button */
  danger?: boolean
  onConfirm: () => void
}

/** Modal yes/no prompt over whatever is open, including the Settings dialog. */
export function ConfirmDialog({
  request,
  cancelLabel,
  onClose,
}: {
  request: ConfirmRequest
  cancelLabel: string
  onClose: () => void
}) {
  const confirmRef = useRef<HTMLButtonElement>(null)

  useEffect(() => {
    confirmRef.current?.focus()
  }, [])

  // capture phase: Escape closes only this prompt, not the Settings dialog under it
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return
      e.stopPropagation()
      onClose()
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose])

  return createPortal(
    <div
      className="cd-overlay"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div
        className="cd-card"
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="cd-title"
        aria-describedby="cd-body"
      >
        <h2 id="cd-title" className="cd-title">
          {request.title}
        </h2>
        <p id="cd-body" className="cd-body">
          {request.body}
        </p>
        <div className="cd-actions">
          <button type="button" className="cd-btn" onClick={onClose}>
            {cancelLabel}
          </button>
          <button
            ref={confirmRef}
            type="button"
            className={`cd-btn ${request.danger ? 'danger' : 'primary'}`}
            onClick={() => {
              onClose()
              request.onConfirm()
            }}
          >
            {request.confirmLabel}
          </button>
        </div>
      </div>
    </div>,
    document.body,
  )
}

/** `[dialog, ask]`: render `dialog` anywhere, call `ask(request)` to prompt. */
export function useConfirm(cancelLabel: string): readonly [ReactNode, (r: ConfirmRequest) => void] {
  const [request, setRequest] = useState<ConfirmRequest | null>(null)
  const close = useCallback(() => setRequest(null), [])
  const dialog = request ? (
    <ConfirmDialog request={request} cancelLabel={cancelLabel} onClose={close} />
  ) : null
  return [dialog, setRequest] as const
}
