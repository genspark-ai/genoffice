import { useEffect, useRef, useState } from 'react'

import { useI18n } from '../i18n/locale'
import { MAX_LABEL_LENGTH, placeholderSource, sanitizeLabel } from '../ai/redact'

interface Props {
  /**
   * What the reader already has for a name. For a single cell that is the
   * cell's own text — the obvious label, offered to be replaced rather than
   * kept. A range has no text of its own, so it starts empty.
   */
  seed: string
  /** A1-style summary of what is about to be withheld, so the scope is never a surprise */
  rangeLabel: string
  onSubmit: (label: string) => void
  onCancel: () => void
}

/**
 * Asks what the withheld cells stand for.
 *
 * The label is the only part of the mark the model ever sees, so the resulting
 * `{{placeholder}}` is shown as it is typed: the reader should know exactly
 * what is about to be exposed before committing. The range is shown for the
 * same reason — a mark can cover a whole column, and that is not a decision
 * anyone should make by accident.
 *
 * It wears the app's own `.modal-backdrop` / `.modal` chrome rather than the
 * markdown app's `.redact-dialog-card`, for the reason the slides dialog
 * records: same behaviour, this app's own CSS.
 *
 * The cells themselves are not touched by this dialog. A mark is a label over
 * the reader's own data, never a replacement for it.
 */
export function RedactDialog({ seed, rangeLabel, onSubmit, onCancel }: Props) {
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
      className="modal-backdrop redact-dialog"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCancel()
      }}
    >
      <div className="modal" role="dialog" aria-modal="true" aria-label={t('redactMenuLabel')}>
        <h2>{t('redactMenuLabel')}</h2>
        <label className="redact-dialog-desc" htmlFor="gx-redact-label">
          {t('redactDialogDesc')}
        </label>
        <p className="redact-dialog-range">{rangeLabel}</p>
        <div className="redact-dialog-preview" aria-live="polite">
          {placeholderSource(label)}
        </div>
        <input
          id="gx-redact-label"
          ref={inputRef}
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
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onCancel}>
            {t('redactCancel')}
          </button>
          <button className="btn-primary" disabled={!label} onClick={() => onSubmit(label)}>
            {t('redactInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}
