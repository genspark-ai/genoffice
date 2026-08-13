import { useState } from 'react'
import type { PivotField } from './PivotDialog'
import { useI18n } from './i18n/locale'

/// Excel's Data → Subtotal, minimal: one "at each change in" column, one
/// aggregated column, SUBTOTAL rows inserted per group plus a grand total.

export interface SubtotalConfig {
  readonly groupCol: number
  readonly valueCol: number
  readonly agg: 'sum' | 'count' | 'average'
}

export function SubtotalDialog({
  fields,
  onCreate,
  onClose,
}: {
  readonly fields: readonly PivotField[]
  /// Returns an error message, or null on success.
  readonly onCreate: (config: SubtotalConfig) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [groupCol, setGroupCol] = useState(fields[0]?.colIndex ?? 0)
  const [valueCol, setValueCol] = useState(fields[1]?.colIndex ?? fields[0]?.colIndex ?? 0)
  const [agg, setAgg] = useState<SubtotalConfig['agg']>('sum')
  const [error, setError] = useState<string | null>(null)

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgSubtotalTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgSubtotalTitle')}</header>
        {fields.length === 0 ? (
          <p className="dialog-note">{t('dlgSubtotalNoFields')}</p>
        ) : (
          <div className="dialog-grid">
            <label>
              {t('dlgSubtotalGroupBy')}
              <select
                value={groupCol}
                onChange={(event) => setGroupCol(Number(event.target.value))}
              >
                {fields.map((field) => (
                  <option key={field.colIndex} value={field.colIndex}>
                    {field.label}
                  </option>
                ))}
              </select>
            </label>
            <label>
              {t('dlgSubtotalFunction')}
              <select
                value={agg}
                onChange={(event) => setAgg(event.target.value as SubtotalConfig['agg'])}
              >
                <option value="sum">{t('dlgPivotAggSum')}</option>
                <option value="count">{t('dlgPivotAggCount')}</option>
                <option value="average">{t('dlgPivotAggAverage')}</option>
              </select>
            </label>
            <label>
              {t('dlgSubtotalAddTo')}
              <select
                value={valueCol}
                onChange={(event) => setValueCol(Number(event.target.value))}
              >
                {fields.map((field) => (
                  <option key={field.colIndex} value={field.colIndex}>
                    {field.label}
                  </option>
                ))}
              </select>
            </label>
            <p className="dialog-note">{t('dlgSubtotalNote')}</p>
          </div>
        )}
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          {fields.length > 0 && (
            <button
              className="primary-action"
              onClick={() => {
                const failure = onCreate({ groupCol, valueCol, agg })
                setError(failure)
                if (failure === null) onClose()
              }}
            >
              {t('dlgOk')}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
