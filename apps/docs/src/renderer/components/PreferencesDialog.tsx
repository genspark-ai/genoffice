/** Word ▸ Preferences ▸ General: measurement units (the only setting so far). */
import { useI18n } from '../i18n/locale'
import { MEASUREMENT_UNITS, setMeasurementUnit, isMeasurementUnit } from '../units'
import { UNIT_NAME_KEYS, useMeasurementUnit } from '../use-measurement'
import { useModalKeys } from './modal-keys'

export function PreferencesDialog({ onClose }: { onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const unit = useMeasurementUnit()

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
        <div className="modal-actions">
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
