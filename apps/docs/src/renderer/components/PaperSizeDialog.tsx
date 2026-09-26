import { useEffect, useRef, useState } from 'react'
import type { SectionSettings } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { useMeasurement } from '../use-measurement'
import { marginsFitPage } from './MarginDialog'
import { LengthInput } from './LengthInput'
import { useModalKeys } from './modal-keys'

/** Word's paper limits: 0.1 in to 22 in */
const MIN_PAGE_TWIPS = 144
const MAX_PAGE_TWIPS = 31680

export type PaperScope = 'section' | 'document'

export function PaperSizeDialog({
  section,
  multiSection,
  onApply,
  onClose,
}: {
  section: SectionSettings
  multiSection: boolean
  /** width/height in the section's current orientation */
  onApply: (width: number, height: number, scope: PaperScope) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const { format } = useMeasurement()
  const [w, setW] = useState(section.pageWidth)
  const [h, setH] = useState(section.pageHeight)
  const [scope, setScope] = useState<PaperScope>(multiSection ? 'section' : 'document')
  // Enter inside a field commits its value first; submit runs on the flushed state
  const [submitTick, setSubmitTick] = useState(0)
  const modalKeys = useModalKeys(onClose)

  const inRange = (v: number) => v >= MIN_PAGE_TWIPS && v <= MAX_PAGE_TWIPS
  const margins = {
    top: section.marginTop,
    right: section.marginRight,
    bottom: section.marginBottom,
    left: section.marginLeft,
  }
  const outOfRange = !inRange(w) || !inRange(h)
  const tooSmall = !outOfRange && !marginsFitPage(margins, w, h)
  const invalid = outOfRange || tooSmall

  const submit = () => {
    if (invalid) return
    onApply(w, h, scope)
    onClose()
  }
  const submitRef = useRef(submit)
  submitRef.current = submit
  useEffect(() => {
    if (submitTick) submitRef.current()
  }, [submitTick])

  // unclamped so an out-of-range entry shows the range message instead of silently snapping
  const field = (label: string, value: number, set: (v: number) => void) => (
    <label>
      {label}
      <LengthInput
        value={value}
        min={0}
        ariaLabel={label}
        live
        onCommit={(twips) => set(twips ?? 0)}
        onEnter={() => setSubmitTick((n) => n + 1)}
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
        <h2>{t('ribbonPaperSizeDialogTitle')}</h2>
        <div className="modal-row margin-row">
          {field(t('ribbonPaperWidth'), w, setW)}
          {field(t('ribbonPaperHeight'), h, setH)}
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonPaperApplyTo')}
            <select value={scope} onChange={(e) => setScope(e.target.value as PaperScope)}>
              <option value="document">{t('ribbonPaperApplyDocument')}</option>
              <option value="section">{t('ribbonPaperApplySection')}</option>
            </select>
          </label>
        </div>
        {outOfRange && (
          <div className="modal-error">
            {t('ribbonPaperSizeOutOfRange', {
              min: format(MIN_PAGE_TWIPS),
              max: format(MAX_PAGE_TWIPS),
            })}
          </div>
        )}
        {tooSmall && <div className="modal-error">{t('ribbonMarginTooLarge')}</div>}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" disabled={invalid} onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
