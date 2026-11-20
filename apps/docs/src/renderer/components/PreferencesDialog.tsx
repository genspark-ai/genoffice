/** Word ▸ Preferences ▸ General: measurement units and the layout-parking
 * opt-in (route-B prototype, #2781 — parking stays off until it is asked for). */
import { useEffect, useState } from 'react'
import { useI18n } from '../i18n/locale'
import { MEASUREMENT_UNITS, setMeasurementUnit, isMeasurementUnit } from '../units'
import { UNIT_NAME_KEYS, useMeasurementUnit } from '../use-measurement'
import {
  layoutParkingOptedIn,
  setLayoutParkingOptedIn,
  subscribeLayoutParkingPref,
} from '../layout-parking-pref'
import { useModalKeys } from './modal-keys'

export function PreferencesDialog({ onClose }: { onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const unit = useMeasurementUnit()
  const [parking, setParking] = useState(layoutParkingOptedIn)
  useEffect(() => subscribeLayoutParkingPref(() => setParking(layoutParkingOptedIn())), [])

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
        <label className="protect-check">
          <input
            type="checkbox"
            checked={parking}
            onChange={(e) => setLayoutParkingOptedIn(e.target.checked)}
          />
          <span className="ctl" aria-hidden="true" />
          {t('appPrefsParkingLabel')}
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
