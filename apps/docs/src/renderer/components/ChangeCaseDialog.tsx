import { useState } from 'react'
import type { CaseMode } from '../editor/case-transform'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

const MODES: Array<[CaseMode, StringKey]> = [
  ['sentence', 'ribbonCaseSentence'],
  ['lower', 'ribbonCaseLower'],
  ['upper', 'ribbonCaseUpper'],
  ['title', 'ribbonCaseTitle'],
  ['toggle', 'ribbonCaseToggle'],
  ['halfWidth', 'ribbonCaseHalfWidth'],
  ['fullWidth', 'ribbonCaseFullWidth'],
]

/** Word's Format > Change Case: one of the cases the Home tab's Change Case menu offers */
export function ChangeCaseDialog({
  onApply,
  onClose,
}: {
  onApply: (mode: CaseMode) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [mode, setMode] = useState<CaseMode>('sentence')
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog case-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          onApply(mode)
        }}
      >
        <h2>{t('ribbonChangeCase')}</h2>
        {MODES.map(([value, labelKey]) => (
          <label key={value} className="tco-check">
            <input
              type="radio"
              name="case"
              checked={mode === value}
              onChange={() => setMode(value)}
            />
            {t(labelKey)}
          </label>
        ))}
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
