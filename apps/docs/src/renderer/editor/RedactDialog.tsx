import { useI18n } from '../i18n/locale'
import { useRedactLabelField } from '@genoffice/ui'

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
 *
 * The words themselves stay in the document either way — withholding a span
 * marks the real text rather than replacing it, so cancelling loses nothing.
 */
export function RedactDialog({ seed, onSubmit, onCancel }: Props) {
  const { t } = useI18n()
  const field = useRedactLabelField(seed, onSubmit, onCancel)

  return (
    <div
      className="modal-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCancel()
      }}
    >
      <div className="modal" role="dialog" aria-modal="true" aria-label={t('redactMenuLabel')}>
        {/* the action is the heading; the explanation is the sentence under it */}
        <h2>{t('redactMenuLabel')}</h2>
        <p className="redact-desc">{t('redactDialogDesc')}</p>
        {/* the claim above is about this app's AI; the file still has the words */}
        <p className="redact-desc">{t('redactDialogScope')}</p>
        <div className="redact-preview">{field.marker}</div>
        <input
          ref={field.inputRef}
          value={field.value}
          maxLength={field.maxLength}
          placeholder={t('redactDialogPlaceholder')}
          onChange={(e) => field.setValue(e.target.value)}
          onKeyDown={field.onKeyDown}
        />
        <div className="modal-actions">
          <button type="button" onClick={onCancel}>
            {t('redactCancel')}
          </button>
          <button
            type="button"
            className="primary"
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
