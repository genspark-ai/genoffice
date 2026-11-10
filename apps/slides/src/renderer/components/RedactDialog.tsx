/**
 * Asks what a withheld span stands for. Reuses the slides .modal-backdrop/.modal
 * conventions (see LinkDialog) rather than the markdown app's own .redact-dialog
 * set — same behaviour, this app's own chrome.
 *
 * The resulting `{{label}}` is shown as it is typed: the label is the only part
 * the model ever sees in place of the words, so the reader should know exactly
 * what is about to be exposed before committing.
 */
import { useRedactLabelField } from '@genoffice/ui'
import { useModalDialog } from './modal-dialog'
import { useI18n } from '../i18n/locale'

interface Props {
  /** the text the span covers, or the shape's name — offered as a starting label */
  seed: string
  onSubmit: (label: string) => void
  onCancel: () => void
}

export function RedactDialog({ seed, onSubmit, onCancel }: Props) {
  const { t } = useI18n()
  const field = useRedactLabelField(seed, onSubmit, onCancel)
  const { titleId, dialogProps } = useModalDialog(onCancel)

  // The seed is a starting point to replace, not a default to keep

  return (
    // data-keep-edit: opening over a live text edit must not commit it — the mark applies to
    // the saved editor selection once the dialog closes
    <div className="modal-backdrop" data-keep-edit="" onClick={onCancel}>
      <div className="modal" {...dialogProps} onClick={(e) => e.stopPropagation()}>
        <h2 id={titleId}>{t('redactMenuLabel')}</h2>
        <p className="modal-note">{t('redactDialogDesc')}</p>
        {/* the claim above is about this app's AI; the file still has the words */}
        <p className="modal-note">{t('redactDialogScope')}</p>
        <div className="redact-preview" aria-live="polite">
          {field.marker}
        </div>
        <input
          ref={field.inputRef}
          type="text"
          value={field.value}
          maxLength={field.maxLength}
          placeholder={t('redactDialogPlaceholder')}
          onChange={(e) => field.setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && field.label) {
              e.preventDefault()
              onSubmit(field.label)
            }
          }}
        />
        <div className="modal-actions">
          <button onClick={onCancel}>{t('redactCancel')}</button>
          <button className="primary" disabled={field.empty} onClick={() => onSubmit(field.label)}>
            {t('redactInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}
