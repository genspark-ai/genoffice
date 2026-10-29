import { useState } from 'react'
import { parseCalibration } from './measure'
import type { TFunc } from './i18n/locale'

/**
 * First-run calibration for the measure tool: the user draws a line over a
 * distance they know (a doorway, a scale bar) and types what it is in the real
 * world. The scale lives in the session only — the file stays geometry-only,
 * and every measurement recomputes from it at commit time.
 */
export function MeasureScaleDialog({
  t,
  drawnPts,
  onApply,
  onClose,
}: {
  t: TFunc
  /** the reference line's PDF length: the ruler the real length is measured against */
  drawnPts: number
  onApply: (unitsPerPt: number, unit: string) => void
  onClose: () => void
}): React.JSX.Element {
  const [value, setValue] = useState('')
  const [unit, setUnit] = useState('m')
  const [error, setError] = useState<string | null>(null)

  const apply = (): void => {
    const scale = parseCalibration(value, unit, drawnPts)
    if (!scale) {
      setError(t('measureScaleErr'))
      return
    }
    onApply(scale.unitsPerPt, scale.unit)
  }

  return (
    <div className="pdf-modal-mask" onClick={onClose}>
      <div
        className="pdf-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('measureScaleTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="pdf-modal-title">{t('measureScaleTitle')}</div>
        <div className="pdf-modal-hint">{t('measureScalePrompt')}</div>
        <div className="pdf-field">
          <label className="pdf-perm-row">
            <span>{t('measureScaleValue')}</span>
            <input
              type="number"
              className="pdf-modal-input"
              value={value}
              autoFocus
              step="any"
              min="0"
              onChange={(e) => {
                setValue(e.target.value)
                setError(null)
              }}
            />
          </label>
          <label className="pdf-perm-row">
            <span>{t('measureScaleUnit')}</span>
            <input
              type="text"
              className="pdf-modal-input"
              value={unit}
              maxLength={12}
              onChange={(e) => {
                setUnit(e.target.value)
                setError(null)
              }}
            />
          </label>
          {error && (
            <div className="pdf-cert-error" role="alert">
              {error}
            </div>
          )}
        </div>
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" onClick={apply}>
            {t('measureScaleApply')}
          </button>
        </div>
      </div>
    </div>
  )
}
