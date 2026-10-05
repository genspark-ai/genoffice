/**
 * Asks what a withheld span stands for. Reuses the slides .modal-backdrop/.modal
 * conventions (see LinkDialog) rather than the markdown app's own .redact-dialog
 * set — same behaviour, this app's own chrome.
 *
 * The resulting `{{label}}` is shown as it is typed: the label is the only part
 * the model ever sees in place of the words, so the reader should know exactly
 * what is about to be exposed before committing.
 */
import { useEffect, useRef, useState } from 'react'
import { useModalDialog } from './modal-dialog'
import { MAX_LABEL_LENGTH, placeholderSource, sanitizeLabel } from '../ai/redact'
import { useI18n } from '../i18n/locale'

interface Props {
  /** the text the span covers, or the shape's name — offered as a starting label */
  seed: string
  onSubmit: (label: string) => void
  onCancel: () => void
}

export function RedactDialog({ seed, onSubmit, onCancel }: Props) {
  const { t } = useI18n()
  const { titleId, dialogProps } = useModalDialog(onCancel)
  const [value, setValue] = useState(seed)
  const inputRef = useRef<HTMLInputElement>(null)

  // The seed is a starting point to replace, not a default to keep
  useEffect(() => inputRef.current?.select(), [])

  const label = sanitizeLabel(value)

  return (
    // data-keep-edit: opening over a live text edit must not commit it — the mark applies to
    // the saved editor selection once the dialog closes
    <div className="modal-backdrop" data-keep-edit="" onClick={onCancel}>
      <div className="modal" {...dialogProps} onClick={(e) => e.stopPropagation()}>
        <h2 id={titleId}>{t('redactMenuLabel')}</h2>
        <p className="modal-note">{t('redactDialogDesc')}</p>
        <div className="redact-preview" aria-live="polite">
          {placeholderSource(label)}
        </div>
        <input
          ref={inputRef}
          type="text"
          value={value}
          maxLength={MAX_LABEL_LENGTH}
          placeholder={t('redactDialogPlaceholder')}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && label) {
              e.preventDefault()
              onSubmit(label)
            }
          }}
        />
        <div className="modal-actions">
          <button onClick={onCancel}>{t('redactCancel')}</button>
          <button className="primary" disabled={!label} onClick={() => onSubmit(label)}>
            {t('redactInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}
