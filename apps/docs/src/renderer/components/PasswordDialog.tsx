/**
 * Password prompt dialog shown when opening a protected docx — either the
 * open password of an ECMA-376 encrypted file (cancel aborts the open) or the
 * password to modify of a write-protected document (cancel opens read-only).
 *
 * Genspark DS form dialog (mockups/protect-dialogs-genspark.html §1/2):
 * top-left title, description line, labelled password field with the shared
 * reveal toggle, error under the field, footer buttons bottom-right.
 * Value / error / busy state live in App's retry loop.
 */
import { FieldError, PasswordInput } from './PasswordInput'

export function PasswordDialog({
  title,
  body,
  label,
  placeholder,
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
  /** label of the password field */
  label: string
  /** placeholder inside the empty field (mockup §1: "enter the open password") */
  placeholder?: string
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
  return (
    <div className="modal-backdrop">
      <div className="modal gs-form pwd-dialog" role="dialog" aria-modal="true">
        <h2>{title}</h2>
        <p className="modal-desc">{body}</p>
        <label className={`fld${error ? ' has-error' : ''}`}>
          {label}
          <PasswordInput
            value={value}
            placeholder={placeholder}
            invalid={!!error}
            autoFocus
            disabled={busy}
            onChange={onChange}
            onKeyDown={(e) => {
              if (e.key === 'Enter') onSubmit()
            }}
          />
          {error && <FieldError>{error}</FieldError>}
        </label>
        <div className="modal-actions">
          <button onClick={onCancel}>{cancelLabel}</button>
          <button className="btn-primary" disabled={busy || !value} onClick={onSubmit}>
            {submitLabel}
          </button>
        </div>
      </div>
    </div>
  )
}
