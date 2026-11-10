/**
 * Design › Slide Size › Custom Size…: width and height in mm, cm or inches
 * (genoffice#2191). The deck rescales through the same setSlideSize op as the
 * presets; each side is held to the 1–56 inch range a .pptx can store.
 */
import { useState } from 'react'
import { useModalDialog } from './modal-dialog'
import { useI18n } from '../i18n/locale'
import { EMU_PER_UNIT, emuToLength, lengthToEmu, type SizeUnit } from '../slide-size'

const UNITS = Object.keys(EMU_PER_UNIT) as SizeUnit[]

export function SlideSizeDialog({
  initial,
  defaultUnit,
  onApply,
  onClose,
}: {
  initial: { cx: number; cy: number }
  defaultUnit: SizeUnit
  onApply: (cx: number, cy: number) => void
  onClose: () => void
}) {
  const { titleId, dialogProps } = useModalDialog(onClose)
  const { t } = useI18n()
  const [unit, setUnit] = useState<SizeUnit>(defaultUnit)
  // `emu` is the exact size behind an untouched field, so unit switches never compound rounding
  const field = (emu: number) => ({
    text: emuToLength(emu, defaultUnit),
    emu: emu as number | null,
  })
  const [width, setWidth] = useState(() => field(initial.cx))
  const [height, setHeight] = useState(() => field(initial.cy))
  const toEmu = (f: { text: string; emu: number | null }) => f.emu ?? lengthToEmu(f.text, unit)
  const cx = toEmu(width)
  const cy = toEmu(height)

  const apply = () => {
    if (cx !== null && cy !== null) onApply(cx, cy)
  }
  // switching units converts what is typed instead of reinterpreting the numbers
  const changeUnit = (next: SizeUnit) => {
    const convert = (f: { text: string; emu: number | null }) => {
      const emu = toEmu(f)
      return emu === null ? f : { text: emuToLength(emu, next), emu }
    }
    setWidth(convert(width))
    setHeight(convert(height))
    setUnit(next)
  }

  const lengthInput = (
    label: string,
    value: { text: string; emu: number | null },
    set: (v: { text: string; emu: number | null }) => void,
    focus = false,
  ) => (
    <label>
      {label}
      <input
        type="text"
        inputMode="decimal"
        value={value.text}
        autoFocus={focus}
        aria-invalid={toEmu(value) === null}
        onChange={(e) => set({ text: e.target.value, emu: null })}
        onKeyDown={(e) => e.key === 'Enter' && apply()}
      />
    </label>
  )

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" {...dialogProps} onClick={(e) => e.stopPropagation()}>
        <h2 id={titleId}>{t('ribbonSlideSizeCustomTitle')}</h2>
        <div className="dlg-two-col">
          {lengthInput(t('ribbonSlideSizeWidth'), width, setWidth, true)}
          {lengthInput(t('ribbonSlideSizeHeight'), height, setHeight)}
        </div>
        <label>
          {t('ribbonSlideSizeUnits')}
          <select value={unit} onChange={(e) => changeUnit(e.target.value as SizeUnit)}>
            {UNITS.map((u) => (
              <option key={u} value={u}>
                {u}
              </option>
            ))}
          </select>
        </label>
        {(cx === null || cy === null) && (
          <p className="dlg-note" role="alert">
            {t('ribbonSlideSizeRange')}
          </p>
        )}
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button className="primary" disabled={cx === null || cy === null} onClick={apply}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
