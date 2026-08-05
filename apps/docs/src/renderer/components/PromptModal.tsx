import { useEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/**
 * In-app replacement for window.prompt, which Electron renderers do not
 * support. Single-line by default; multiline renders a textarea (notes).
 */
export function PromptModal({
  title,
  placeholder,
  initial = '',
  multiline = false,
  allowEmpty = false,
  onSubmit,
  onClose,
}: {
  title: string
  placeholder?: string
  initial?: string
  multiline?: boolean
  /** submit button stays enabled for empty input (e.g. clearing a value) */
  allowEmpty?: boolean
  onSubmit: (value: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [value, setValue] = useState(initial)
  const inputRef = useRef<HTMLInputElement>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)
  const modalKeys = useModalKeys(onClose)

  useEffect(() => {
    const el = multiline ? textareaRef.current : inputRef.current
    el?.focus()
    el?.select()
  }, [multiline])

  const submit = () => {
    if (!allowEmpty && !value.trim()) return
    onSubmit(value.trim())
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{title}</h2>
        <label>
          {multiline ? (
            <textarea
              ref={textareaRef}
              className="prompt-textarea"
              rows={4}
              value={value}
              placeholder={placeholder}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit()
              }}
            />
          ) : (
            <input
              ref={inputRef}
              value={value}
              placeholder={placeholder}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit()
              }}
            />
          )}
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" disabled={!allowEmpty && !value.trim()} onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
