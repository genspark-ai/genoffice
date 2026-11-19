import { useState } from 'react'
import { useI18n } from '../i18n/locale'
import { useMeasurement } from '../use-measurement'
import { useModalKeys } from './modal-keys'

/** Word's Format > Asian Layout > Fit Text: the width the selected text is spread across */
export function FitTextDialog({
  width,
  onApply,
  onClose,
}: {
  /** the width the text is fitted into now, in twips; 0 when it is not fitted */
  width: number
  /** 0 takes the fitting off */
  onApply: (twips: number) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const measure = useMeasurement()
  const keys = useModalKeys(onClose)
  const [text, setText] = useState(width > 0 ? measure.format(width) : '')
  const typed = measure.parse(text)
  const twips = typed !== null && typed > 0 ? Math.round(typed) : null
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog fit-text-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (twips === null) return
          onClose()
          if (twips !== width) onApply(twips)
        }}
      >
        <h2>{t('ribbonFitText')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonFitWidth')}
            <input
              type="text"
              value={text}
              autoFocus
              spellCheck={false}
              onChange={(e) => setText(e.target.value)}
            />
          </label>
        </div>
        <div className="modal-actions">
          <button
            type="button"
            disabled={width <= 0}
            onClick={() => {
              onClose()
              onApply(0)
            }}
          >
            {t('ribbonFitRemove')}
          </button>
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={twips === null}>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
