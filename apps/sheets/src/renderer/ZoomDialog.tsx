import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'
import { clampZoomPercent, SHEET_ZOOM_MAX, SHEET_ZOOM_MIN } from './zoom-range'

const PRESETS = [200, 100, 75, 50, 25] as const
type Choice = (typeof PRESETS)[number] | 'fit' | 'custom'

export function ZoomDialog({
  zoomPercent,
  onCommand,
  onClose,
}: {
  readonly zoomPercent: number
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const preset = PRESETS.find((value) => value === zoomPercent)
  const [choice, setChoice] = useState<Choice>(preset ?? 'custom')
  const [custom, setCustom] = useState(String(zoomPercent))
  const [error, setError] = useState(false)
  const modal = useModalDialog(onClose)

  const apply = (): void => {
    if (choice === 'fit') {
      onCommand('zoom-to-selection')
    } else if (choice === 'custom') {
      const value = Number(custom.trim().replace(/%$/, ''))
      if (!Number.isFinite(value) || value !== clampZoomPercent(value)) {
        setError(true)
        return
      }
      onCommand(`zoom:${value}`)
    } else {
      onCommand(`zoom:${choice}`)
    }
    onClose()
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        {...modal}
        className="format-cells-dialog zoom-dialog"
        role="dialog"
        aria-label={t('dlgZoomTitle')}
        onClick={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          modal.onKeyDown(event)
          // A focused button (Cancel) keeps its own Enter activation.
          if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement)) {
            event.preventDefault()
            apply()
          }
        }}
      >
        <header>{t('dlgZoomTitle')}</header>
        <section className="dialog-body">
          <fieldset className="zoom-choices">
            <legend>{t('dlgZoomMagnification')}</legend>
            {PRESETS.map((value) => (
              <label key={value} className="dialog-check">
                <input
                  type="radio"
                  name="zoom-choice"
                  autoFocus={choice === value}
                  checked={choice === value}
                  onChange={() => setChoice(value)}
                />
                {value}%
              </label>
            ))}
            <label className="dialog-check">
              <input
                type="radio"
                name="zoom-choice"
                checked={choice === 'fit'}
                onChange={() => setChoice('fit')}
              />
              {t('dlgZoomFitSelection')}
            </label>
            <label className="dialog-check zoom-custom">
              <input
                type="radio"
                name="zoom-choice"
                checked={choice === 'custom'}
                onChange={() => setChoice('custom')}
              />
              {t('dlgZoomCustom')}
              <input
                type="number"
                min={SHEET_ZOOM_MIN}
                max={SHEET_ZOOM_MAX}
                autoFocus={choice === 'custom'}
                value={custom}
                onFocus={() => setChoice('custom')}
                onChange={(event) => {
                  setCustom(event.target.value)
                  setError(false)
                }}
              />
              %
            </label>
          </fieldset>
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {t('dlgZoomRangeError', { min: SHEET_ZOOM_MIN, max: SHEET_ZOOM_MAX })}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" onClick={apply}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
