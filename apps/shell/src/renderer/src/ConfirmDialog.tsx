import { useId, useRef } from 'react'
import { createPortal } from 'react-dom'
import { useModalFocus } from './use-modal-focus'

/**
 * Yes/no prompt in the shared `.modal` style, layered over whatever is open
 * (including Settings). Tab stays inside it, focus returns to the opener on
 * close, and Escape closes only this prompt. A `danger` prompt opens on
 * Cancel so Enter alone never runs the destructive action.
 */
export function ConfirmDialog({
  title,
  body,
  confirmLabel,
  cancelLabel,
  danger = false,
  onConfirm,
  onCancel,
}: {
  title: string
  body: string
  confirmLabel: string
  cancelLabel: string
  danger?: boolean
  onConfirm: () => void
  onCancel: () => void
}) {
  const cardRef = useRef<HTMLDivElement>(null)
  const cancelRef = useRef<HTMLButtonElement>(null)
  const confirmRef = useRef<HTMLButtonElement>(null)
  const titleId = useId()
  const bodyId = useId()
  useModalFocus(cardRef, danger ? cancelRef : confirmRef, onCancel)

  return createPortal(
    <div className="modal-overlay" onClick={onCancel}>
      <div
        ref={cardRef}
        className="modal"
        role="alertdialog"
        aria-modal="true"
        aria-labelledby={titleId}
        aria-describedby={bodyId}
        onClick={(e) => e.stopPropagation()}
      >
        <h3 id={titleId}>{title}</h3>
        <p id={bodyId}>{body}</p>
        <div className="modal-buttons">
          <button ref={cancelRef} type="button" className="btn btn-secondary" onClick={onCancel}>
            {cancelLabel}
          </button>
          <button
            ref={confirmRef}
            type="button"
            className={`btn ${danger ? 'btn-danger' : 'btn-primary'}`}
            onClick={onConfirm}
          >
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>,
    document.body,
  )
}
