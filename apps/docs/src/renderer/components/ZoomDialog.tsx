import { useState } from 'react'
import { useI18n } from '../i18n/locale'
import { clampDocsZoom, DOCS_ZOOM_MAX, DOCS_ZOOM_MIN, type ZoomFitMode } from '../wheel-zoom'
import { useModalKeys } from './modal-keys'

const PRESETS = [200, 100, 75] as const
type Preset = (typeof PRESETS)[number]
type Choice = Preset | ZoomFitMode | 'custom'

interface ZoomDialogProps {
  zoom: number
  /** the zoom a fit mode would land on right now, for the Percent readout */
  fitValue: (mode: ZoomFitMode) => number | null
  onApply: (target: number | ZoomFitMode) => void
  onClose: () => void
}

function isPreset(n: number): n is Preset {
  return (PRESETS as readonly number[]).includes(n)
}

/** Word's View ▸ Zoom dialog: preset radios, fit modes and a 10–500 percent box. */
export function ZoomDialog({ zoom, fitValue, onApply, onClose }: ZoomDialogProps) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const start = Math.round(zoom)
  const [choice, setChoice] = useState<Choice>(isPreset(start) ? start : 'custom')
  const [percent, setPercent] = useState(String(start))

  const pick = (c: Choice) => {
    setChoice(c)
    if (typeof c === 'number') setPercent(String(c))
    else if (c !== 'custom') {
      const v = fitValue(c)
      if (v != null) setPercent(String(v))
    }
  }
  const commitPercent = () => {
    const typed = percent.trim() === '' ? NaN : Number(percent)
    const n = clampDocsZoom(Number.isFinite(typed) ? Math.round(typed) : start)
    setPercent(String(n))
    return n
  }
  const submit = () => {
    if (typeof choice === 'string' && choice !== 'custom') onApply(choice)
    else onApply(commitPercent())
    onClose()
  }
  const radio = (c: Choice, label: string) => (
    <label className="zoom-dialog-radio">
      <input type="radio" name="zoom-to" checked={choice === c} onChange={() => pick(c)} />
      {label}
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form zoom-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          submit()
        }}
      >
        <h2>{t('ribbonZoomDialog')}</h2>
        <fieldset className="zoom-dialog-group">
          <legend>{t('ribbonZoomTo')}</legend>
          <div className="zoom-dialog-radios">
            {radio(200, '200%')}
            {radio('width', t('ribbonPageWidth'))}
            {radio(100, '100%')}
            {radio('text', t('ribbonZoomTextWidth'))}
            {radio(75, '75%')}
            {radio('page', t('ribbonWholePage'))}
          </div>
        </fieldset>
        <label className="zoom-dialog-percent">
          <span>{t('ribbonZoomPercent')}</span>
          <input
            type="number"
            min={DOCS_ZOOM_MIN}
            max={DOCS_ZOOM_MAX}
            step={1}
            value={percent}
            onChange={(e) => {
              setPercent(e.target.value)
              const n = Number(e.target.value)
              setChoice(isPreset(n) ? n : 'custom')
            }}
            onBlur={commitPercent}
          />
          <span>%</span>
        </label>
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="submit" className="primary">
            {t('ribbonOk')}
          </button>
        </div>
      </form>
    </div>
  )
}
