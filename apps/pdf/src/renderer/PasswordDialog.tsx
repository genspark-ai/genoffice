/**
 * Password prompt dialog shown when opening an encrypted PDF, matching the
 * docs PasswordDialog (mockups/protect-dialogs-genspark.html §1/2): top-left
 * title, description with the quoted file name, labelled password field with
 * a reveal toggle, inline error, cancel / open footer buttons.
 *
 * Value / wrong state live in App's retry loop (status === 'password').
 */
import { useLayoutEffect, useRef, useState } from 'react'
import { useI18n } from './i18n/locale'
import { IconAlert, IconEye, IconEyeOff } from './icons'

export function PasswordDialog({
  fileName,
  value,
  wrong,
  onChange,
  onSubmit,
  onCancel,
}: {
  fileName: string
  value: string
  /** the previous attempt failed: styles the field and shakes it */
  wrong: boolean
  onChange: (value: string) => void
  onSubmit: () => void
  onCancel: () => void
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
    <div className="pdf-modal-mask pdf-pwd-mask">
      <div className="pdf-modal pdf-pwd-dialog" role="dialog" aria-modal="true">
        <h2>{t('pwTitle')}</h2>
        <p className="pdf-pwd-desc">{t('pwBody', { name: fileName })}</p>
        <label className={`pdf-pwd-fld${wrong ? ' has-error' : ''}`}>
          {t('pwLabel')}
          <span className="pdf-pwd-wrap">
            <input
              ref={inputRef}
              type={show ? 'text' : 'password'}
              className={wrong ? 'invalid' : undefined}
              autoFocus
              placeholder={t('pwPlaceholder')}
              value={value}
              onChange={(e) => onChange(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && value) onSubmit()
              }}
            />
            <button
              type="button"
              className="pdf-pwd-eye"
              tabIndex={-1}
              aria-label={t(show ? 'pwHide' : 'pwShow')}
              // keep the caret in the password field: without this, clicking the
              // toggle focuses the button and further typing goes nowhere
              onMouseDown={(e) => e.preventDefault()}
              onClick={toggleShow}
            >
              {show ? <IconEyeOff size={16} /> : <IconEye size={16} />}
            </button>
          </span>
          {wrong && (
            <span className="pdf-pwd-error" role="alert">
              <IconAlert size={13} />
              {t('pwWrong')}
            </span>
          )}
        </label>
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onCancel}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" disabled={!value} onClick={onSubmit}>
            {t('pwOpen')}
          </button>
        </div>
      </div>
    </div>
  )
}
