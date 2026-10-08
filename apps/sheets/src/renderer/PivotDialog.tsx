import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

/// Create PivotTable: source range (from the selection) and destination cell.
/// The layout itself is edited in the PivotTable Fields pane that opens once
/// the pivot exists, like Excel.

export type {
  OoXmlPivotConfig,
  PivotField,
  PivotGroupingOption,
  PivotLabelFilterOption,
  PivotValueFilterOption,
  PivotValueSpec,
} from './pivot-field-model'

export interface PivotCreateRequest {
  readonly sourceRange: string
  readonly targetCell: string
}

export function PivotDialog({
  hasFields,
  sourceRange,
  onCreate,
  onClose,
}: {
  readonly hasFields: boolean
  readonly sourceRange: string
  /// Returns an error message, or null on success.
  readonly onCreate: (request: PivotCreateRequest) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [targetCell, setTargetCell] = useState('')
  const [error, setError] = useState<string | null>(null)
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog pivot-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgPivotCreateTitle')}
        style={{ minWidth: 380 }}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgPivotCreateTitle')}</header>
        {!hasFields ? (
          <p className="dialog-note">{t('dlgPivotNoFields')}</p>
        ) : (
          <div className="dialog-grid">
            <label>
              {t('dlgPivotSourceRange')}
              <input type="text" className="cell-input" value={sourceRange} readOnly />
            </label>
            <label>
              {t('dlgPivotTargetCell')}
              <input
                type="text"
                className="cell-input"
                placeholder={t('dlgPivotTargetPlaceholder')}
                value={targetCell}
                onChange={(event) => setTargetCell(event.target.value.toUpperCase())}
                style={{ width: 80 }}
              />
            </label>
            <p className="dialog-note dialog-span">{t('dlgPivotCreateHint')}</p>
          </div>
        )}
        {error && (
          <p className="dialog-note dialog-error" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          {hasFields && (
            <button
              className="primary-action"
              onClick={() => {
                const failure = onCreate({ sourceRange, targetCell: targetCell.trim() || 'A1' })
                if (failure !== null) setError(failure)
                else onClose()
              }}
            >
              {t('dlgCreate')}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
