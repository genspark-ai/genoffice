import { useRedactLabelField } from '@genoffice/ui'

import { useI18n } from '../i18n/locale'

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
  const field = useRedactLabelField(seed, onSubmit, onCancel)

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
          {/* the claim above is about this app's AI; the file still has the words */}
          <p className="redact-dialog-scope">{t('redactDialogScope')}</p>
        </label>
        <p className="redact-dialog-range">{rangeLabel}</p>
        <div className="redact-dialog-preview" aria-live="polite">
          {field.marker}
        </div>
        <input
          id="gx-redact-label"
          ref={field.inputRef}
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
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onCancel}>
            {t('redactCancel')}
          </button>
          <button
            className="btn-primary"
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
