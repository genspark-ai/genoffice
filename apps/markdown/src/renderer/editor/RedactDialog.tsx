import { useRedactLabelField } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'

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
  const field = useRedactLabelField(seed, onSubmit, onCancel)

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
        {/* the claim above is about this app's AI; the file still has the words */}
        <p className="redact-dialog-scope">{t('redactDialogScope')}</p>
        <div className="redact-dialog-preview">{field.marker}</div>
        <input
          ref={field.inputRef}
          className="redact-dialog-input"
          value={field.value}
          maxLength={field.maxLength}
          placeholder={t('redactDialogPlaceholder')}
          onChange={(e) => field.setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && field.label) {
              e.preventDefault()
              onSubmit(field.label)
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
            disabled={field.empty}
            onClick={() => onSubmit(field.label)}
          >
            {t('redactInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}
