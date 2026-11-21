import { useState } from 'react'
import type { TabStop } from '@genoffice/docx-engine'
import { useI18n, type StringKey } from '../i18n/locale'
import { useMeasurement } from '../use-measurement'
import { LengthInput } from './LengthInput'
import { useModalKeys } from './modal-keys'
import { withTabStop } from './tab-stops'

type Align = Exclude<TabStop['val'], 'clear'>
type Leader = NonNullable<TabStop['leader']>

const ALIGNS: Array<[Align, StringKey]> = [
  ['left', 'appTabLeft'],
  ['center', 'appTabCenter'],
  ['right', 'appTabRight'],
  ['decimal', 'appTabDecimal'],
  ['bar', 'appTabBar'],
]
// the leaders Word's dialog offers, drawn as the characters they fill the gap with
const LEADERS: Array<[Leader, string]> = [
  ['dot', '.......'],
  ['hyphen', '-------'],
  ['underscore', '_______'],
]

// Word's range for the default tab stop: 0.01" to 22"
const MIN_DEFAULT_STOP_TWIPS = 14
const MAX_DEFAULT_STOP_TWIPS = 31680

/** Word's Format > Tabs: the paragraph's tab stops, each with its alignment and leader */
export function TabsDialog({
  stops: initial,
  defaultStop,
  onApply,
  onClose,
}: {
  /** the stops the paragraph has, its style's among them */
  stops: TabStop[]
  /** the document's default tab stop, in twips */
  defaultStop: number
  /** the paragraph's stops, and the document's default tab stop in twips */
  onApply: (stops: TabStop[], defaultStop: number) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const measure = useMeasurement()
  const keys = useModalKeys(onClose)
  const [stops, setStops] = useState(() => initial.filter((s) => s.val !== 'clear'))
  const [position, setPosition] = useState('')
  const [align, setAlign] = useState<Align>('left')
  const [leader, setLeader] = useState<Leader>('none')
  const [defaultTabStop, setDefaultTabStop] = useState(defaultStop)
  // a listed stop keeps its own position: its text, read back, may round to another twip
  const [shown, setShown] = useState<number | null>(null)
  const typed = measure.parse(position)
  const typedPos =
    shown !== null && position === measure.format(shown)
      ? shown
      : typed !== null && typed > 0
        ? Math.round(typed)
        : null
  const picked = stops.find((s) => s.pos === typedPos)

  const show = (stop: TabStop) => {
    setShown(stop.pos)
    setPosition(measure.format(stop.pos))
    setAlign(stop.val === 'clear' ? 'left' : stop.val)
    setLeader(stop.leader ?? 'none')
  }
  /** the stops with the one the fields describe set; unchanged while the position is not a length */
  const withTyped = (): TabStop[] =>
    typedPos === null
      ? stops
      : withTabStop(stops, {
          pos: typedPos,
          val: align,
          ...(leader === 'none' ? {} : { leader }),
        })

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog tabs-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          // OK also sets a stop typed and not yet set, as Word does
          onApply(withTyped(), defaultTabStop)
        }}
      >
        <h2>{t('ribbonTabsTitle')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonTabsPosition')}
            <input
              type="text"
              value={position}
              spellCheck={false}
              onChange={(e) => setPosition(e.target.value)}
            />
          </label>
        </div>
        <div className="tabs-dialog-list" role="listbox" aria-label={t('ribbonTabsTitle')}>
          {stops.map((stop) => (
            <button
              key={stop.pos}
              type="button"
              role="option"
              aria-selected={stop === picked}
              className={stop === picked ? 'active' : ''}
              onClick={() => show(stop)}
            >
              {measure.format(stop.pos)}
            </button>
          ))}
        </div>
        <div className="doc-dialog-caption">{t('appAlignment')}</div>
        <div className="tabs-dialog-options">
          {ALIGNS.map(([value, labelKey]) => (
            <label key={value} className="tco-check">
              <input
                type="radio"
                name="tab-align"
                checked={align === value}
                onChange={() => setAlign(value)}
              />
              {t(labelKey)}
            </label>
          ))}
        </div>
        <div className="doc-dialog-caption">{t('ribbonTabsLeader')}</div>
        <div className="tabs-dialog-options">
          <label className="tco-check">
            <input
              type="radio"
              name="tab-leader"
              checked={leader === 'none'}
              onChange={() => setLeader('none')}
            />
            {t('ribbonDropCapNone')}
          </label>
          {LEADERS.map(([value, sample]) => (
            <label key={value} className="tco-check">
              <input
                type="radio"
                name="tab-leader"
                checked={leader === value}
                onChange={() => setLeader(value)}
              />
              {sample}
            </label>
          ))}
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonTabsDefault')}
            <LengthInput
              value={defaultTabStop}
              min={MIN_DEFAULT_STOP_TWIPS}
              max={MAX_DEFAULT_STOP_TWIPS}
              ariaLabel={t('ribbonTabsDefault')}
              live
              onCommit={(twips) => twips !== null && setDefaultTabStop(twips)}
            />
          </label>
        </div>
        <div className="modal-actions">
          <div className="tabs-dialog-edit">
            <button
              type="button"
              disabled={typedPos === null}
              onClick={() => setStops(withTyped())}
            >
              {t('ribbonTabsSet')}
            </button>
            <button
              type="button"
              disabled={!picked}
              onClick={() => {
                setStops(stops.filter((s) => s !== picked))
                setPosition('')
              }}
            >
              {t('appTabClear')}
            </button>
            <button
              type="button"
              disabled={stops.length === 0}
              onClick={() => {
                setStops([])
                setPosition('')
              }}
            >
              {t('ribbonClearAll')}
            </button>
          </div>
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
