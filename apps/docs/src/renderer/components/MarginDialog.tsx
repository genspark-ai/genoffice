import { useState } from 'react'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

export interface PageMargins {
  top: number
  right: number
  bottom: number
  left: number
}

const TWIPS_PER_CM = 1440 / 2.54
/** Word rejects margins that leave less than about one inch of body */
const MIN_BODY_TWIPS = 1440

export const cmFromTwips = (twips: number): number => Math.round((twips / TWIPS_PER_CM) * 100) / 100

export const twipsFromCmInput = (value: string, originalTwips: number): number => {
  if (value === String(cmFromTwips(originalTwips))) return originalTwips
  return Math.round(Math.max(0, Number(value) || 0) * TWIPS_PER_CM)
}

export const marginsFitPage = (
  margins: PageMargins,
  pageWidth: number,
  pageHeight: number,
): boolean =>
  margins.left + margins.right <= pageWidth - MIN_BODY_TWIPS &&
  margins.top + margins.bottom <= pageHeight - MIN_BODY_TWIPS

type Side = 'top' | 'bottom' | 'left' | 'right'

const SIDES: Array<[Side, StringKey]> = [
  ['top', 'ribbonMarginTop'],
  ['bottom', 'ribbonMarginBottom'],
  ['left', 'ribbonMarginLeft'],
  ['right', 'ribbonMarginRight'],
]

export function MarginDialog({
  margins,
  pageWidth,
  pageHeight,
  onApply,
  onClose,
}: {
  margins: PageMargins
  pageWidth: number
  pageHeight: number
  onApply: (next: PageMargins) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [values, setValues] = useState<Record<Side, string>>({
    top: String(cmFromTwips(margins.top)),
    bottom: String(cmFromTwips(margins.bottom)),
    left: String(cmFromTwips(margins.left)),
    right: String(cmFromTwips(margins.right)),
  })
  const modalKeys = useModalKeys(onClose)

  const twips = (side: Side) => twipsFromCmInput(values[side], margins[side])
  const next: PageMargins = {
    top: twips('top'),
    right: twips('right'),
    bottom: twips('bottom'),
    left: twips('left'),
  }
  const tooLarge = !marginsFitPage(next, pageWidth, pageHeight)

  const submit = () => {
    if (tooLarge) return
    onApply(next)
    onClose()
  }

  const field = ([side, labelKey]: [Side, StringKey]) => (
    <label key={side}>
      {t(labelKey)} (cm)
      <input
        type="number"
        min={0}
        step={0.1}
        value={values[side]}
        onChange={(e) => setValues((v) => ({ ...v, [side]: e.target.value }))}
        onKeyDown={(e) => {
          if (e.key === 'Enter') submit()
        }}
      />
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{t('ribbonMarginDialogTitle')}</h2>
        <div className="modal-row margin-row">{SIDES.slice(0, 2).map(field)}</div>
        <div className="modal-row margin-row">{SIDES.slice(2).map(field)}</div>
        {tooLarge && <div className="modal-error">{t('ribbonMarginTooLarge')}</div>}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" disabled={tooLarge} onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
