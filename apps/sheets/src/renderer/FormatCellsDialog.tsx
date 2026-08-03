import { useEffect, useRef, useState } from 'react'

import { draftFromSelection, formatCellsCommands, type FormatCellsDraft } from './format-cells'
import { useI18n, type StringKey } from './i18n/locale'

import type { SelectionFormat } from './selection-format'

/// Excel's Format Cells dialog (⌘1), scoped to what the save pipeline can
/// persist today: number format, alignment, font, border, fill, and
/// protection flags. It opens prefilled with the selection's current format
/// and emits the same command strings the ribbon uses — but only for
/// settings the user actually changed.

const TABS = ['Number', 'Alignment', 'Font', 'Border', 'Fill', 'Protection'] as const
type Tab = (typeof TABS)[number]

const TAB_LABELS: Record<Tab, StringKey> = {
  Number: 'dlgFcTabNumber',
  Alignment: 'dlgFcTabAlignment',
  Font: 'dlgFcTabFont',
  Border: 'dlgFcTabBorder',
  Fill: 'dlgFcTabFill',
  Protection: 'dlgFcTabProtection',
}

const H_ALIGNMENTS = ['left', 'center', 'right', 'justify', 'distributed'] as const
const V_ALIGNMENTS = ['top', 'middle', 'bottom'] as const
const H_ALIGN_LABELS: Record<(typeof H_ALIGNMENTS)[number], StringKey> = {
  left: 'dlgFcAlignLeft',
  center: 'dlgFcAlignCenter',
  right: 'dlgFcAlignRight',
  justify: 'dlgFcAlignJustify',
  distributed: 'dlgFcAlignDistributed',
}
const V_ALIGN_LABELS: Record<(typeof V_ALIGNMENTS)[number], StringKey> = {
  top: 'dlgFcAlignTop',
  middle: 'dlgFcAlignMiddle',
  bottom: 'dlgFcAlignBottom',
}
const ROTATIONS: { readonly labelKey: StringKey; readonly value: string }[] = [
  { labelKey: 'dlgFcRotNone', value: '0' },
  { labelKey: 'dlgFcRotCcw', value: '45' },
  { labelKey: 'dlgFcRotCw', value: '-45' },
  { labelKey: 'dlgFcRotUp', value: '90' },
  { labelKey: 'dlgFcRotDown', value: '-90' },
  { labelKey: 'dlgFcRotVertical', value: 'vertical' },
]

const NUMBER_CATEGORIES: { readonly labelKey: StringKey; readonly pattern: string }[] = [
  { labelKey: 'dlgFcNumGeneral', pattern: 'General' },
  { labelKey: 'dlgFcNumNumber', pattern: '0.00' },
  { labelKey: 'dlgFcNumThousands', pattern: '#,##0.00' },
  { labelKey: 'dlgFcNumCurrency', pattern: '$#,##0.00' },
  { labelKey: 'dlgFcNumCurrencyCny', pattern: '¥#,##0.00' },
  { labelKey: 'dlgFcNumPercent', pattern: '0.00%' },
  { labelKey: 'dlgFcNumDate', pattern: 'yyyy-mm-dd' },
  { labelKey: 'dlgFcNumTime', pattern: 'hh:mm:ss' },
  { labelKey: 'dlgFcNumScientific', pattern: '0.00E+00' },
  { labelKey: 'dlgFcNumText', pattern: '@' },
]

const FONT_FAMILIES = ['Aptos', 'Arial', 'Calibri', 'Times New Roman', '微软雅黑', '宋体']
const FONT_SIZES = ['9', '10', '11', '12', '14', '16', '18', '22', '26']
/// OOXML ST_BorderStyle names offered by the line-style picker; the
/// journal maps them 1:1 to Univer BorderStyleTypes and back to the file.
const BORDER_LINE_STYLES: { readonly labelKey: StringKey; readonly value: string }[] = [
  { labelKey: 'dlgFcStyleThin', value: 'thin' },
  { labelKey: 'dlgFcStyleMedium', value: 'medium' },
  { labelKey: 'dlgFcStyleThick', value: 'thick' },
  { labelKey: 'dlgFcStyleDouble', value: 'double' },
  { labelKey: 'dlgFcStyleHair', value: 'hair' },
  { labelKey: 'dlgFcStyleDashed', value: 'dashed' },
  { labelKey: 'dlgFcStyleDotted', value: 'dotted' },
]

const BORDER_PRESETS: { readonly labelKey: StringKey; readonly value: string }[] = [
  { labelKey: 'dlgFcBorderNone', value: 'none' },
  { labelKey: 'dlgFcBorderAll', value: 'all' },
  { labelKey: 'dlgFcBorderOutline', value: 'outer' },
  { labelKey: 'dlgFcBorderThickOutline', value: 'thick-outer' },
  { labelKey: 'dlgFcBorderTop', value: 'top' },
  { labelKey: 'dlgFcBorderBottom', value: 'bottom' },
  { labelKey: 'dlgFcBorderLeft', value: 'left' },
  { labelKey: 'dlgFcBorderRight', value: 'right' },
]

export function FormatCellsDialog({
  selectionFormat,
  onCommand,
  onClose,
}: {
  readonly selectionFormat: SelectionFormat | null
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [tab, setTab] = useState<Tab>('Number')
  useEffect(() => {
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])
  // The selection format at open time is the change baseline; the live echo
  // must not move it while the dialog is up.
  const initialRef = useRef(draftFromSelection(selectionFormat))
  const initial = initialRef.current
  const [draft, setDraft] = useState(initial)
  const set = <K extends keyof FormatCellsDraft>(key: K, value: FormatCellsDraft[K]): void =>
    setDraft((previous) => ({ ...previous, [key]: value }))

  const namedCategories = NUMBER_CATEGORIES.map((category) => ({
    name: t(category.labelKey),
    pattern: category.pattern,
  }))
  const categoryOptions =
    !draft.pattern || NUMBER_CATEGORIES.some((category) => category.pattern === draft.pattern)
      ? namedCategories
      : [
          ...namedCategories,
          { name: t('dlgFcNumCustom', { pattern: draft.pattern }), pattern: draft.pattern },
        ]
  const familyOptions =
    !draft.family || FONT_FAMILIES.includes(draft.family)
      ? FONT_FAMILIES
      : [draft.family, ...FONT_FAMILIES]
  const sizeOptions =
    !draft.size || FONT_SIZES.includes(draft.size)
      ? FONT_SIZES
      : [...FONT_SIZES, draft.size].sort((a, b) => Number(a) - Number(b))

  function handleApply(): void {
    for (const command of formatCellsCommands(initial, draft)) onCommand(command)
    onClose()
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgFcTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgFcTitle')}</header>
        <nav className="dialog-tabs">
          {TABS.map((candidate) => (
            <button
              key={candidate}
              className={candidate === tab ? 'active' : ''}
              onClick={() => setTab(candidate)}
            >
              {t(TAB_LABELS[candidate])}
            </button>
          ))}
        </nav>
        <section className="dialog-body">
          {tab === 'Number' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcCategory')}
                <select value={draft.pattern} onChange={(e) => set('pattern', e.target.value)}>
                  <option value="">{t('dlgFcUnchanged')}</option>
                  {categoryOptions.map((category) => (
                    <option key={category.pattern} value={category.pattern}>
                      {category.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t('dlgFcCustomCode')}
                <input
                  value={draft.pattern}
                  placeholder="#,##0.00"
                  onChange={(e) => set('pattern', e.target.value)}
                />
              </label>
            </div>
          )}
          {tab === 'Alignment' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcHorizontal')}
                <select value={draft.hAlign} onChange={(e) => set('hAlign', e.target.value)}>
                  <option value="">{t('dlgFcUnchanged')}</option>
                  {H_ALIGNMENTS.map((h) => (
                    <option key={h} value={h}>
                      {t(H_ALIGN_LABELS[h])}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t('dlgFcVertical')}
                <select value={draft.vAlign} onChange={(e) => set('vAlign', e.target.value)}>
                  <option value="">{t('dlgFcUnchanged')}</option>
                  {V_ALIGNMENTS.map((v) => (
                    <option key={v} value={v}>
                      {t(V_ALIGN_LABELS[v])}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t('dlgFcWrapText')}
                <select
                  value={draft.wrapText}
                  onChange={(e) => set('wrapText', e.target.value as FormatCellsDraft['wrapText'])}
                >
                  <option value="">{t('dlgFcUnchanged')}</option>
                  <option value="on">{t('dlgFcOn')}</option>
                  <option value="off">{t('dlgFcOff')}</option>
                </select>
              </label>
              <label>
                {t('dlgFcOrientation')}
                <select value={draft.rotation} onChange={(e) => set('rotation', e.target.value)}>
                  <option value="">{t('dlgFcUnchanged')}</option>
                  {ROTATIONS.map((r) => (
                    <option key={r.value} value={r.value}>
                      {t(r.labelKey)}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t('dlgFcIndent')}
                <input
                  type="number"
                  min="0"
                  max="250"
                  value={draft.indent}
                  placeholder={t('dlgFcUnchanged')}
                  onChange={(e) => set('indent', e.target.value)}
                />
              </label>
            </div>
          )}
          {tab === 'Font' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcFont')}
                <select value={draft.family} onChange={(e) => set('family', e.target.value)}>
                  <option value="">{t('dlgFcUnchanged')}</option>
                  {familyOptions.map((f) => (
                    <option key={f}>{f}</option>
                  ))}
                </select>
              </label>
              <label>
                {t('dlgFcSize')}
                <select value={draft.size} onChange={(e) => set('size', e.target.value)}>
                  <option value="">{t('dlgFcUnchanged')}</option>
                  {sizeOptions.map((s) => (
                    <option key={s}>{s}</option>
                  ))}
                </select>
              </label>
              {(
                [
                  ['dlgFcBold', 'bold'],
                  ['dlgFcItalic', 'italic'],
                  ['dlgFcUnderline', 'underline'],
                  ['dlgFcStrikethrough', 'strike'],
                ] as const
              ).map(([labelKey, key]) => (
                <label key={key}>
                  {t(labelKey)}
                  <select
                    value={draft[key]}
                    onChange={(e) => set(key, e.target.value as FormatCellsDraft[typeof key])}
                  >
                    <option value="">{t('dlgFcUnchanged')}</option>
                    <option value="on">{t('dlgFcOn')}</option>
                    <option value="off">{t('dlgFcOff')}</option>
                  </select>
                </label>
              ))}
              <label>
                {t('dlgFcColor')}
                <input
                  type="color"
                  value={draft.fontColor || '#000000'}
                  onChange={(e) => set('fontColor', e.target.value)}
                />
              </label>
            </div>
          )}
          {tab === 'Border' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcBorderPresets')}
                <select value={draft.border} onChange={(e) => set('border', e.target.value)}>
                  <option value="">{t('dlgFcUnchanged')}</option>
                  {BORDER_PRESETS.map((preset) => (
                    <option key={preset.value} value={preset.value}>
                      {t(preset.labelKey)}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t('dlgFcBorderStyle')}
                <select
                  value={draft.borderStyle}
                  onChange={(e) => set('borderStyle', e.target.value)}
                >
                  {BORDER_LINE_STYLES.map((style) => (
                    <option key={style.value} value={style.value}>
                      {t(style.labelKey)}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t('dlgFcColor')}
                <input
                  type="color"
                  value={draft.borderColor}
                  onChange={(e) => set('borderColor', e.target.value)}
                />
              </label>
            </div>
          )}
          {tab === 'Fill' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcBackground')}
                <input
                  type="color"
                  value={draft.fill || '#ffffff'}
                  disabled={draft.noFill}
                  // An unset fill already displays as white, so picking white
                  // never fires a change event. Seed the draft when the user
                  // opens the picker: that turns the displayed white into an
                  // explicit choice, distinguishable from "no fill".
                  onClick={() => {
                    if (!draft.fill) set('fill', '#ffffff')
                  }}
                  onChange={(e) => set('fill', e.target.value)}
                />
              </label>
              <label className="dialog-check">
                <input
                  type="checkbox"
                  checked={draft.noFill}
                  onChange={(e) => set('noFill', e.target.checked)}
                />
                {t('dlgFcNoFill')}
              </label>
            </div>
          )}
          {tab === 'Protection' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcLocked')}
                <select
                  value={draft.locked}
                  onChange={(e) => set('locked', e.target.value as FormatCellsDraft['locked'])}
                >
                  <option value="">{t('dlgFcUnchanged')}</option>
                  <option value="on">{t('dlgFcOn')}</option>
                  <option value="off">{t('dlgFcOff')}</option>
                </select>
              </label>
              <label>
                {t('dlgFcHidden')}
                <select
                  value={draft.hidden}
                  onChange={(e) => set('hidden', e.target.value as FormatCellsDraft['hidden'])}
                >
                  <option value="">{t('dlgFcUnchanged')}</option>
                  <option value="on">{t('dlgFcOn')}</option>
                  <option value="off">{t('dlgFcOff')}</option>
                </select>
              </label>
              <p className="dialog-note">{t('dlgFcProtectionNote')}</p>
            </div>
          )}
        </section>
        <footer className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" onClick={handleApply}>
            {t('dlgOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
