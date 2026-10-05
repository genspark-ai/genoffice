import { useEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import { MAX_LABEL_LENGTH, placeholderSource, sanitizeLabel } from './redact'

interface Props {
  /** the text the span covers, offered as a starting label */
  seed: string
  onSubmit: (label: string) => void
  onCancel: () => void
}

/**
 * Asks what the withheld span stands for. The label is the only part the model
 * ever sees, so the dialog shows the resulting marker as it is typed: the
 * reader should know exactly what is about to be exposed before committing.
 */
export function RedactDialog({ seed, onSubmit, onCancel }: Props) {
  const { t } = useI18n()
  const [value, setValue] = useState(seed)
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
  }, [])

  const label = sanitizeLabel(value)

  return (
    <div
      className="redact-dialog"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCancel()
      }}
    >
      <div
        className="redact-dialog-card"
        role="dialog"
        aria-modal="true"
        aria-label={t('redactMenuLabel')}
      >
        {/* the action is the heading; the explanation is the sentence under it */}
        <h2 className="redact-dialog-title">{t('redactMenuLabel')}</h2>
        <p className="redact-dialog-desc">{t('redactDialogDesc')}</p>
        <div className="redact-dialog-preview">{placeholderSource(label)}</div>
        <input
          ref={inputRef}
          className="redact-dialog-input"
          value={value}
          maxLength={MAX_LABEL_LENGTH}
          placeholder={t('redactDialogPlaceholder')}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && label) {
              e.preventDefault()
              onSubmit(label)
            } else if (e.key === 'Escape') {
              e.preventDefault()
              onCancel()
            }
          }}
        />
        <div className="redact-dialog-row">
          <button type="button" className="redact-dialog-btn" onClick={onCancel}>
            {t('redactCancel')}
          </button>
          <button
            type="button"
            className="redact-dialog-btn primary"
            disabled={!label}
            onClick={() => onSubmit(label)}
          >
            {t('redactInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}
