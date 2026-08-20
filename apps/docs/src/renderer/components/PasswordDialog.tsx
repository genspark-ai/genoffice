/**
 * Password prompt dialog shown when opening a protected docx — either the
 * open password of an ECMA-376 encrypted file (cancel aborts the open) or the
 * password to modify of a write-protected document (cancel opens read-only).
 *
 * Purely presentational: value / error / busy state live in App's retry loop.
 * All texts arrive already translated so the two prompts can share one look.
 */
import { useLayoutEffect, useRef, useState } from 'react'
import { IconAlert, IconEye, IconEyeOff, IconLock } from './icons'

export function PasswordDialog({
  title,
  body,
  label,
  showLabel,
  hideLabel,
  value,
  error,
  busy = false,
  submitLabel,
  cancelLabel,
  onChange,
  onSubmit,
  onCancel,
}: {
  title: string
  /** explanatory line including the quoted document name */
  body: string
  /** placeholder / accessible name of the password field */
  label: string
  /** accessible names of the visibility toggle's two states */
  showLabel: string
  hideLabel: string
  value: string
  /** translated failure line ('' = none); styles the field and shakes it */
  error: string
  /** a decrypt attempt is in flight: field and submit lock up */
  busy?: boolean
  submitLabel: string
  cancelLabel: string
  onChange: (value: string) => void
  onSubmit: () => void
  onCancel: () => void
}) {
  const [show, setShow] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  /** caret to restore after the type=password/text swap resets it to 0 */
  const caretRef = useRef<[number, number] | null>(null)

  const toggleShow = () => {
    const el = inputRef.current
    caretRef.current = el
      ? [el.selectionStart ?? value.length, el.selectionEnd ?? value.length]
      : null
    setShow((s) => !s)
  }

  useLayoutEffect(() => {
    const el = inputRef.current
    const caret = caretRef.current
    caretRef.current = null
    if (!el || !caret) return
    const restore = () => {
      if (document.activeElement === el) el.setSelectionRange(caret[0], caret[1])
    }
    restore()
    // Chromium resets the caret again asynchronously after the type swap
    const raf = requestAnimationFrame(restore)
    return () => cancelAnimationFrame(raf)
  }, [show])

  return (
    <div className="modal-backdrop">
      <div className="modal pwd-dialog" role="dialog" aria-modal="true">
        <div className="pwd-dialog-badge">
          <IconLock size={26} />
        </div>
        <h2>{title}</h2>
        <p className="pwd-dialog-body">{body}</p>
        <div className={`pwd-field${error ? ' has-error' : ''}`}>
          <input
            ref={inputRef}
            type={show ? 'text' : 'password'}
            autoFocus
            placeholder={label}
            aria-label={label}
            disabled={busy}
            value={value}
            onChange={(e) => onChange(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') onSubmit()
            }}
          />
          <button
            type="button"
            className="pwd-eye"
            tabIndex={-1}
            aria-label={show ? hideLabel : showLabel}
            disabled={busy}
            // keep the caret in the password field: without this, clicking the
            // toggle focuses the button and further typing goes nowhere
            onMouseDown={(e) => e.preventDefault()}
            onClick={toggleShow}
          >
            {show ? <IconEyeOff size={16} /> : <IconEye size={16} />}
          </button>
        </div>
        {error && (
          <p className="modal-error pwd-dialog-error">
            <IconAlert size={14} />
            <span>{error}</span>
          </p>
        )}
        <div className="modal-actions pwd-dialog-actions">
          <button onClick={onCancel}>{cancelLabel}</button>
          <button className="btn-primary" disabled={busy || !value} onClick={onSubmit}>
            {submitLabel}
          </button>
        </div>
      </div>
    </div>
  )
}
