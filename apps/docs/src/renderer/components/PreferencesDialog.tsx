/** Word ▸ Preferences: measurement units (General) and the style area width (View). */
import { useI18n } from '../i18n/locale'
import { setStyleAreaTwips, STYLE_AREA_MAX_TWIPS } from '../style-area'
import {
  MEASUREMENT_UNITS,
  UNIT_STEP,
  UNIT_SUFFIX,
  fromUnit,
  isMeasurementUnit,
  setMeasurementUnit,
  toUnit,
} from '../units'
import { UNIT_NAME_KEYS, useMeasurementUnit, useStyleAreaTwips } from '../use-measurement'
import { useModalKeys } from './modal-keys'

export function PreferencesDialog({ onClose }: { onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const unit = useMeasurementUnit()
  const styleArea = useStyleAreaTwips()

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal protect-dialog" role="dialog" aria-label={t('appPrefsTitle')}>
        <h2>{t('appPrefsTitle')}</h2>
        <label className="fld prefs-field">
          <span>{t('appPrefsUnitsLabel')}</span>
          <select
            value={unit}
            aria-label={t('appPrefsUnitsLabel')}
            onChange={(e) => {
              if (isMeasurementUnit(e.target.value)) setMeasurementUnit(e.target.value)
            }}
          >
            {MEASUREMENT_UNITS.map((u) => (
              <option key={u} value={u}>
                {t(UNIT_NAME_KEYS[u])}
              </option>
            ))}
          </select>
        </label>
        <label className="fld prefs-field">
          <span>{t('appPrefsStyleArea')}</span>
          <span className="prefs-length">
            <input
              type="number"
              className="len-input"
              min={0}
              max={toUnit(STYLE_AREA_MAX_TWIPS, unit)}
              step={UNIT_STEP[unit]}
              value={toUnit(styleArea, unit)}
              aria-label={t('appPrefsStyleArea')}
              onChange={(e) => {
                const value = Number(e.target.value)
                if (Number.isFinite(value)) setStyleAreaTwips(fromUnit(value, unit))
              }}
            />
            <span className="para-unit">{UNIT_SUFFIX[unit]}</span>
          </span>
        </label>
        <div className="modal-actions">
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
