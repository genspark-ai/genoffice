/**
 * Password form field with a show/hide (eye) reveal toggle, for dialogs with
 * several password fields (ProtectDialog). The single-password prompt dialogs
 * use the standalone PasswordDialog instead.
 *
 * `hideReveal` removes the toggle: used while a field still holds the
 * untypable KEEP sentinel that stands in for an existing hash-only password,
 * where revealing would only show garbage.
 */
import { useLayoutEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import { IconAlert, IconEye, IconEyeOff } from './icons'

export function PasswordInput({
  value,
  onChange,
  onKeyDown,
  autoFocus,
  disabled,
  maxLength,
  invalid,
  hideReveal,
}: {
  value: string
  onChange: (value: string) => void
  onKeyDown?: (e: React.KeyboardEvent<HTMLInputElement>) => void
  autoFocus?: boolean
  disabled?: boolean
  maxLength?: number
  invalid?: boolean
  hideReveal?: boolean
}) {
  const { t } = useI18n()
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
    <span className="pwd-wrap">
      <input
        ref={inputRef}
        type={show && !hideReveal ? 'text' : 'password'}
        className={invalid ? 'invalid' : undefined}
        autoFocus={autoFocus}
        disabled={disabled}
        maxLength={maxLength}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={onKeyDown}
      />
      {!hideReveal && (
        <button
          type="button"
          className="pwd-eye"
          tabIndex={-1}
          disabled={disabled}
          aria-label={t(show ? 'appPwdHide' : 'appPwdShow')}
          // keep the caret in the password field: without this, clicking the
          // toggle focuses the button and further typing goes nowhere
          onMouseDown={(e) => e.preventDefault()}
          onClick={toggleShow}
        >
          {show ? <IconEyeOff size={16} /> : <IconEye size={16} />}
        </button>
      )}
    </span>
  )
}

/** field-level validation message: alert icon + text, announced via role=alert */
export function FieldError({ children }: { children: React.ReactNode }) {
  return (
    <span className="fld-error" role="alert">
      <IconAlert size={13} />
      {children}
    </span>
  )
}
