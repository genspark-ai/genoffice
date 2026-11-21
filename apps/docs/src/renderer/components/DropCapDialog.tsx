import { useMemo, useState } from 'react'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import { fontFamiliesFor } from '../font-list'
import { LengthInput } from './LengthInput'
import {
  DROP_CAP_DEFAULT_LINES,
  DROP_CAP_MAX_GAP,
  DROP_CAP_MAX_LINES,
  clampDropCapLines,
  dropCapGapOf,
  dropCapLinesOf,
} from '../editor/drop-cap'

export type DropCapPosition = 'none' | 'drop' | 'margin'
export interface DropCapValue {
  position: DropCapPosition
  lines: number
  /** distance from the text, twips */
  gap: number
}

/** the paragraph's dropCap attribute (JSON of type and lines) as the dialog shows it */
export function dropCapValueOf(attr: unknown): DropCapValue {
  const none: DropCapValue = { position: 'none', lines: DROP_CAP_DEFAULT_LINES, gap: 0 }
  if (typeof attr !== 'string') return none
  try {
    const parsed = JSON.parse(attr) as { type?: string; lines?: number }
    if (parsed.type !== 'drop' && parsed.type !== 'margin') return none
    return { position: parsed.type, lines: dropCapLinesOf(attr), gap: dropCapGapOf(attr) }
  } catch {
    return none
  }
}

const POSITIONS: Array<[DropCapPosition, StringKey]> = [
  ['none', 'ribbonDropCapNone'],
  ['drop', 'ribbonDropCapDropped'],
  ['margin', 'ribbonDropCapMargin'],
]

/** Word's Format > Drop Cap: where the first letter sits and how many lines it spans */
export function DropCapDialog({
  current,
  font,
  onApply,
  onClose,
}: {
  current: DropCapValue
  /** the font of the paragraph's first letter; '' when it cannot be told */
  font: string
  /** `font` is passed only when another was picked */
  onApply: (value: DropCapValue, font?: string) => void
  onClose: () => void
}) {
  const { t, lang } = useI18n()
  const keys = useModalKeys(onClose)
  const [position, setPosition] = useState(current.position)
  const [lines, setLines] = useState(String(current.lines))
  const [gap, setGap] = useState(current.gap)
  const [face, setFace] = useState(font)
  const fonts = useMemo(() => {
    const list = [...fontFamiliesFor(lang)]
    return font && !list.includes(font) ? [font, ...list] : list
  }, [lang, font])
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog dropcap-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          onApply(
            { position, lines: clampDropCapLines(lines), gap },
            position !== 'none' && face && face !== font ? face : undefined,
          )
        }}
      >
        <h2>{t('ribbonDropCap')}</h2>
        <div className="doc-dialog-caption">{t('ribbonDropCapPosition')}</div>
        {POSITIONS.map(([value, labelKey]) => (
          <label key={value} className="tco-check">
            <input
              type="radio"
              name="dropcap"
              checked={position === value}
              onChange={() => setPosition(value)}
            />
            {t(labelKey)}
          </label>
        ))}
        <div className="modal-row margin-row">
          <label>
            {t('appFontTabFont')}
            <select
              name="font"
              value={face}
              disabled={position === 'none'}
              onChange={(e) => setFace(e.target.value)}
            >
              {!font && <option value="" />}
              {fonts.map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonDropCapLines')}
            <input
              type="number"
              min={1}
              max={DROP_CAP_MAX_LINES}
              value={lines}
              disabled={position === 'none'}
              onChange={(e) => setLines(e.target.value)}
            />
          </label>
          <label>
            {t('ribbonDropCapDistance')}
            <LengthInput
              value={gap}
              max={DROP_CAP_MAX_GAP}
              live
              disabled={position === 'none'}
              onCommit={(twips) => setGap(twips ?? 0)}
            />
          </label>
        </div>
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
