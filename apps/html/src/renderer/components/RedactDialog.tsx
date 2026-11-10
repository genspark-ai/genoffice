import { useId } from 'react'
import { useRedactLabelField } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'

/**
 * Asks what a withheld span stands for, and shows the marker that results from
 * the answer as it is typed.
 *
 * The live preview is the point of the dialog. The label is the only part of a
 * withheld span a model ever sees, so the reader should know exactly what is
 * about to stand in for their words before they commit to it.
 *
 * Chrome is this app's own — a mask and a card, the shape of the image dialogs
 * this app already opens — rather than the slides app's `.modal-backdrop` set,
 * which does not exist in this stylesheet.
 */
interface Props {
  /** the selected text, offered as a starting label */
  seed: string
  onSubmit: (label: string) => void
  onCancel: () => void
}

export function RedactDialog({ seed, onSubmit, onCancel }: Props) {
  const { t } = useI18n()
  const field = useRedactLabelField(seed, onSubmit, onCancel)
  const titleId = useId()

  // The seed is a starting point to replace, not a default to keep

  return (
    <div
      className="redact-dialog-mask"
      onClick={onCancel}
      onKeyDown={(e) => {
        // Escape is handled here rather than by a document listener so the app's
        // own shortcuts see a dialog that is already closing
        if (e.key === 'Escape') {
          e.stopPropagation()
          onCancel()
        }
      }}
    >
      <div
        className="redact-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        onClick={(e) => e.stopPropagation()}
      >
        <h2 className="redact-dialog-title" id={titleId}>
          {t('redactMenuLabel')}
        </h2>
        <p className="redact-dialog-desc">{t('redactDialogDesc')}</p>
        {/* the claim above is about this app's AI; the file still has the words */}
        <p className="redact-dialog-scope">{t('redactDialogScope')}</p>
        <div className="redact-dialog-preview" aria-live="polite">
          {field.marker}
        </div>
        <input
          ref={field.inputRef}
          className="redact-dialog-input"
          type="text"
          value={field.value}
          maxLength={field.maxLength}
          autoFocus
          placeholder={t('redactDialogPlaceholder')}
          onChange={(e) => field.setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && field.label) {
              e.preventDefault()
              onSubmit(field.label)
            }
          }}
        />
        <div className="redact-dialog-actions">
          <button className="redact-dialog-btn" onClick={onCancel}>
            {t('redactCancel')}
          </button>
          <button
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
