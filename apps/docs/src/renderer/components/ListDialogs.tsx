import { useMemo, useState } from 'react'
import type { ChangeEvent, ReactNode } from 'react'
import { Dropdown } from '@genoffice/ui'
import { formatNumber, type CustomNumberingLevel, type StyleInfo } from '@genoffice/docx-engine'
import { useI18n, type StringKey } from '../i18n/locale'
import { fontFamiliesFor } from '../font-list'
import { useSystemFontFamilies } from '../system-fonts'
import { useModalKeys } from './modal-keys'
import { LengthInput } from './LengthInput'
import {
  LIST_NUM_FMTS,
  MULTILEVEL_LIBRARY,
  bulletPresetLevels,
  numFmtSample,
  numberPresetLevels,
  previewLevelText,
} from '../list-presets'

export type ListDialogKind = 'bullet' | 'number' | 'multi' | 'value' | 'indents'

type Suff = NonNullable<CustomNumberingLevel['suff']>
type Jc = NonNullable<CustomNumberingLevel['lvlJc']>

const SUFF_OPTIONS: Array<{ value: Suff; labelKey: StringKey }> = [
  { value: 'tab', labelKey: 'ribbonFollowTab' },
  { value: 'space', labelKey: 'ribbonFollowSpace' },
  { value: 'nothing', labelKey: 'ribbonFollowNothing' },
]
const JC_OPTIONS: Array<{ value: Jc; labelKey: StringKey }> = [
  { value: 'left', labelKey: 'appAlignLeft' },
  { value: 'center', labelKey: 'appAlignCenter' },
  { value: 'right', labelKey: 'appAlignRight' },
]

/** Unicode bullets grouped the way Word's Symbol dialog subsets read */
const BULLET_SYMBOLS =
  '•◦▪▫■□●○◆◇▶▷►‣⁃‐–—' + '→⇒➢➤➔➜➡✓✔✗✘★☆✦✧✩✪✳' + '❀❈❍❖❤♥♦♣♠☺☀☁☂☎✉✎✑✒' + '†‡§¶°»›◈◉◎⬤⬛⬜⬥'

/** list geometry fields: Word bounds them by the page (0–22") */
const MAX_LIST_TWIPS = 31680

function Modal({
  title,
  className,
  onClose,
  onOk,
  okDisabled,
  children,
}: {
  title: string
  className: string
  onClose: () => void
  onOk: () => void
  okDisabled?: boolean
  children: ReactNode
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className={`modal list-dialog ${className}`} role="dialog" aria-label={title}>
        <h2>{title}</h2>
        {children}
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="button" className="primary" onClick={onOk} disabled={okDisabled}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

function NumFmtDropdown({
  value,
  withBullet,
  onPick,
}: {
  value: string
  withBullet?: boolean
  onPick: (numFmt: string) => void
}) {
  const { t } = useI18n()
  const options = useMemo(() => {
    const fmts: string[] = withBullet ? ['bullet', ...LIST_NUM_FMTS] : [...LIST_NUM_FMTS]
    if (!fmts.includes(value)) fmts.unshift(value)
    return fmts.map((f) => ({ value: f, label: numFmtSample(f) }))
  }, [value, withBullet])
  return (
    <Dropdown
      value={value}
      ariaLabel={t('ribbonListNumStyle')}
      options={options}
      onPick={onPick}
      className="list-dialog-fmt"
    />
  )
}

/** Marker font / size / color / weight (the level's w:rPr); empty = follow the text */
function MarkerFontFields({
  level,
  onChange,
}: {
  level: CustomNumberingLevel
  onChange: (patch: Partial<CustomNumberingLevel>) => void
}) {
  const { t, lang } = useI18n()
  const { families: systemFamilies, load } = useSystemFontFamilies()
  const fontOptions = useMemo(() => {
    const names = new Set<string>([...fontFamiliesFor(lang), ...systemFamilies])
    if (level.font) names.add(level.font)
    return [
      { value: '', label: t('ribbonAuto') },
      ...[...names].map((f) => ({ value: f, label: f })),
    ]
  }, [lang, systemFamilies, level.font, t])
  return (
    <div className="list-dialog-font" onMouseEnter={load}>
      <label className="fld">
        <span>{t('ribbonBulletFont')}</span>
        <Dropdown
          value={level.font ?? ''}
          ariaLabel={t('ribbonBulletFont')}
          options={fontOptions}
          onPick={(font) => onChange({ font: font || undefined })}
          className="list-dialog-fontname"
        />
      </label>
      <label className="fld list-dialog-num">
        <span>{t('ribbonMarkerSizePt')}</span>
        <input
          type="number"
          min={1}
          max={400}
          step={0.5}
          placeholder={t('ribbonAuto')}
          value={level.szHalfPoints ? level.szHalfPoints / 2 : ''}
          onChange={(e) => {
            const v = parseFloat(e.target.value)
            onChange({ szHalfPoints: Number.isFinite(v) && v > 0 ? Math.round(v * 2) : undefined })
          }}
        />
      </label>
      <label className="fld list-dialog-num">
        <span>{t('ribbonMarkerColor')}</span>
        <span className="list-dialog-color">
          <input
            type="color"
            aria-label={t('ribbonMarkerColor')}
            value={`#${level.color ?? '000000'}`}
            onChange={(e) => onChange({ color: e.target.value.slice(1).toUpperCase() })}
          />
          <button
            type="button"
            className="list-dialog-link"
            disabled={!level.color}
            onClick={() => onChange({ color: undefined })}
          >
            {t('ribbonAuto')}
          </button>
        </span>
      </label>
      <label className="font-check">
        <input
          type="checkbox"
          checked={!!level.bold}
          onChange={(e) => onChange({ bold: e.target.checked || undefined })}
        />
        <span>{t('ribbonMarkerBold')}</span>
      </label>
      <label className="font-check">
        <input
          type="checkbox"
          checked={!!level.italic}
          onChange={(e) => onChange({ italic: e.target.checked || undefined })}
        />
        <span>{t('ribbonMarkerItalic')}</span>
      </label>
    </div>
  )
}

function LevelPreview({
  levels,
  picSrc,
  count = 3,
}: {
  levels: CustomNumberingLevel[]
  picSrc?: string
  count?: number
}) {
  const { t } = useI18n()
  const first = levels[0]
  const style = {
    fontFamily: first?.font,
    color: first?.color ? `#${first.color}` : undefined,
    fontWeight: first?.bold ? 700 : undefined,
    fontStyle: first?.italic ? 'italic' : undefined,
    fontSize: first?.szHalfPoints ? `${first.szHalfPoints / 2}pt` : undefined,
  }
  return (
    <div className="list-dialog-preview" aria-label={t('ribbonListPreviewCol')}>
      {Array.from({ length: count }, (_, row) => {
        const value = (first?.start ?? 1) + row
        const text =
          first?.numFmt === 'bullet'
            ? first.lvlText
            : (first?.lvlText ?? '').replace(/%1/g, formatNumber(value, first?.numFmt ?? 'decimal'))
        return (
          <div key={row} className="list-dialog-preview-row">
            {picSrc ? (
              <img src={picSrc} alt="" className="list-dialog-preview-pic" />
            ) : (
              <span className="list-dialog-preview-marker" style={style}>
                {text}
              </span>
            )}
            <span className="list-gallery-preview-line" />
          </div>
        )
      })}
    </div>
  )
}

/* ================= Set Numbering Value ================= */

export function SetNumberingValueDialog({
  numFmt,
  lvlText,
  currentValue,
  canContinue,
  onApply,
  onClose,
}: {
  numFmt: string
  lvlText: string
  /** the number the caret's item shows now */
  currentValue: number
  canContinue: boolean
  onApply: (mode: 'restart' | 'continue', value: number | undefined) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [mode, setMode] = useState<'restart' | 'continue'>('restart')
  const [advance, setAdvance] = useState(false)
  const [value, setValue] = useState(currentValue)
  const valueEnabled = mode === 'restart' || advance
  const preview = lvlText.replace(/%\d/g, formatNumber(valueEnabled ? value : currentValue, numFmt))
  return (
    <Modal
      title={t('ribbonSetNumberingValue')}
      className="list-value-dialog"
      onClose={onClose}
      onOk={() => onApply(mode, valueEnabled ? value : undefined)}
    >
      <label className="font-check">
        <input
          type="radio"
          name="list-value-mode"
          checked={mode === 'restart'}
          onChange={() => setMode('restart')}
        />
        <span>{t('ribbonStartNewList')}</span>
      </label>
      <label className="font-check">
        <input
          type="radio"
          name="list-value-mode"
          disabled={!canContinue}
          checked={mode === 'continue'}
          onChange={() => setMode('continue')}
        />
        <span>{t('ribbonContinuePreviousList')}</span>
      </label>
      <label className="font-check list-dialog-indent">
        <input
          type="checkbox"
          disabled={mode !== 'continue'}
          checked={mode === 'continue' && advance}
          onChange={(e) => setAdvance(e.target.checked)}
        />
        <span>{t('ribbonAdvanceValue')}</span>
      </label>
      <label className="fld list-dialog-num">
        <span>{t('ribbonSetValueTo')}</span>
        <input
          type="number"
          min={0}
          max={32767}
          disabled={!valueEnabled}
          value={value}
          onChange={(e) => setValue(Math.max(0, Math.floor(Number(e.target.value) || 0)))}
        />
      </label>
      <div className="list-dialog-preview-line">
        <span>{t('ribbonListPreviewCol')}:</span>
        <strong>{preview}</strong>
      </div>
    </Modal>
  )
}

/* ================= Adjust List Indents ================= */

export function AdjustListIndentsDialog({
  level,
  onApply,
  onClose,
}: {
  level: CustomNumberingLevel
  onApply: (patch: Partial<CustomNumberingLevel>) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const hanging = level.hanging ?? 0
  const [markerTwips, setMarkerTwips] = useState(Math.max(0, level.indentLeft - hanging))
  const [textTwips, setTextTwips] = useState(level.indentLeft)
  const [suff, setSuff] = useState<Suff>(level.suff ?? 'tab')
  const [addTab, setAddTab] = useState(level.tabStop !== undefined)
  const [tabTwips, setTabTwips] = useState(level.tabStop ?? level.indentLeft)
  const apply = () => {
    const indentLeft = textTwips
    const marker = Math.min(markerTwips, indentLeft)
    onApply({
      indentLeft,
      hanging: indentLeft - marker,
      suff,
      tabStop: suff === 'tab' && addTab ? tabTwips : undefined,
    })
  }
  const num = (value: number, set: (v: number) => void, label: string) => (
    <label className="fld list-dialog-num">
      <span>{label}</span>
      <LengthInput
        value={value}
        min={0}
        max={MAX_LIST_TWIPS}
        ariaLabel={label}
        live
        onCommit={(twips) => set(twips ?? 0)}
      />
    </label>
  )
  return (
    <Modal
      title={t('ribbonAdjustListIndents')}
      className="list-indents-dialog"
      onClose={onClose}
      onOk={apply}
    >
      {num(markerTwips, setMarkerTwips, t('ribbonNumberPosition'))}
      {num(textTwips, setTextTwips, t('ribbonTextIndent'))}
      <label className="fld">
        <span>{t('ribbonFollowNumberWith')}</span>
        <Dropdown
          value={suff}
          ariaLabel={t('ribbonFollowNumberWith')}
          options={SUFF_OPTIONS.map((o) => ({ value: o.value, label: t(o.labelKey) }))}
          onPick={setSuff}
        />
      </label>
      <label className="font-check">
        <input
          type="checkbox"
          disabled={suff !== 'tab'}
          checked={suff === 'tab' && addTab}
          onChange={(e) => setAddTab(e.target.checked)}
        />
        <span>{t('ribbonAddTabStopAt')}</span>
      </label>
      {suff === 'tab' && addTab && num(tabTwips, setTabTwips, t('ribbonAddTabStopAt'))}
    </Modal>
  )
}

/* ================= Define New Bullet ================= */

export interface BulletPicture {
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
}

/** the marker formatting a define dialog owns; pattern, glyph and geometry come from the preset helpers */
function markerExtras(level: CustomNumberingLevel): Partial<CustomNumberingLevel> {
  const out: Partial<CustomNumberingLevel> = {}
  if (level.font) out.font = level.font
  if (level.szHalfPoints) out.szHalfPoints = level.szHalfPoints
  if (level.color) out.color = level.color
  if (level.bold) out.bold = true
  if (level.italic) out.italic = true
  if (level.lvlJc && level.lvlJc !== 'left') out.lvlJc = level.lvlJc
  return out
}

export function DefineBulletDialog({
  ilvl = 0,
  onApply,
  onClose,
}: {
  /** the caret's list level: Word's Define New Bullet changes that level */
  ilvl?: number
  /** the picture stays in dialog state until OK, so Cancel leaves nothing behind */
  onApply: (levels: CustomNumberingLevel[], picture: BulletPicture | null) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [level, setLevel] = useState<CustomNumberingLevel>({
    numFmt: 'bullet',
    lvlText: '\u2022',
    indentLeft: 720,
    hanging: 360,
    lvlJc: 'left',
  })
  const [picture, setPicture] = useState<BulletPicture | null>(null)
  const [tab, setTab] = useState<'symbol' | 'picture'>('symbol')
  const patch = (p: Partial<CustomNumberingLevel>) => setLevel((l) => ({ ...l, ...p }))
  const onFile = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    const mime = file.type
    if (mime !== 'image/png' && mime !== 'image/jpeg' && mime !== 'image/gif') return
    const bytes = new Uint8Array(await file.arrayBuffer())
    let bin = ''
    for (const b of bytes) bin += String.fromCharCode(b)
    setPicture({ base64: btoa(bin), mime })
    setTab('picture')
  }
  const apply = () => {
    const usePicture = tab === 'picture' && picture !== null
    onApply(
      bulletPresetLevels(level.lvlText, markerExtras(level), ilvl),
      usePicture ? picture : null,
    )
  }
  return (
    <Modal
      title={t('ribbonDefineNewBullet')}
      className="list-bullet-dialog"
      onClose={onClose}
      onOk={apply}
      okDisabled={tab === 'picture' && picture === null}
    >
      <div className="dlg-tabs" role="tablist">
        {(['symbol', 'picture'] as const).map((k) => (
          <button
            key={k}
            type="button"
            role="tab"
            aria-selected={tab === k}
            className={tab === k ? 'active' : undefined}
            onClick={() => setTab(k)}
          >
            {t(k === 'symbol' ? 'ribbonBulletSymbol' : 'ribbonBulletPicture')}
          </button>
        ))}
      </div>
      {tab === 'symbol' ? (
        <>
          <div
            className="list-dialog-symbols"
            role="listbox"
            aria-label={t('ribbonBulletCharacter')}
          >
            {[...BULLET_SYMBOLS].map((ch) => (
              <button
                key={ch}
                type="button"
                role="option"
                aria-selected={level.lvlText === ch}
                className={level.lvlText === ch ? 'selected' : undefined}
                style={{ fontFamily: level.font }}
                onClick={() => patch({ lvlText: ch })}
              >
                {ch}
              </button>
            ))}
          </div>
          <label className="fld list-dialog-num">
            <span>{t('ribbonBulletCharacter')}</span>
            <input
              value={level.lvlText}
              maxLength={2}
              onChange={(e) => patch({ lvlText: [...e.target.value].slice(0, 1).join('') || '•' })}
            />
          </label>
          <MarkerFontFields level={level} onChange={patch} />
        </>
      ) : (
        <div className="list-dialog-picture">
          <label className="list-dialog-file">
            <input type="file" accept="image/png,image/jpeg,image/gif" onChange={onFile} />
            <span className="list-dialog-file-button">{t('ribbonChooseImage')}</span>
          </label>
          {picture !== null && (
            <span className="list-dialog-hint">{t('ribbonPictureBulletChosen')}</span>
          )}
        </div>
      )}
      <label className="fld">
        <span>{t('ribbonListAlignment')}</span>
        <Dropdown
          value={level.lvlJc ?? 'left'}
          ariaLabel={t('ribbonListAlignment')}
          options={JC_OPTIONS.map((o) => ({ value: o.value, label: t(o.labelKey) }))}
          onPick={(lvlJc) => patch({ lvlJc })}
        />
      </label>
      <LevelPreview
        levels={[level]}
        picSrc={
          tab === 'picture' && picture ? `data:${picture.mime};base64,${picture.base64}` : undefined
        }
      />
    </Modal>
  )
}

/* ================= Define New Number Format ================= */

export function DefineNumberFormatDialog({
  ilvl = 0,
  initial,
  onApply,
  onClose,
}: {
  /** the caret's list level: the defined format lands there */
  ilvl?: number
  initial?: CustomNumberingLevel
  onApply: (levels: CustomNumberingLevel[]) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [level, setLevel] = useState<CustomNumberingLevel>(
    initial ?? { numFmt: 'decimal', lvlText: '%1.', indentLeft: 720, hanging: 360, start: 1 },
  )
  const patch = (p: Partial<CustomNumberingLevel>) => setLevel((l) => ({ ...l, ...p }))
  return (
    <Modal
      title={t('ribbonDefineNewNumber')}
      className="list-number-dialog"
      onClose={onClose}
      onOk={() =>
        onApply(
          numberPresetLevels(
            level.numFmt,
            level.lvlText,
            { ...markerExtras(level), start: level.start ?? 1 },
            ilvl,
          ),
        )
      }
    >
      <label className="fld">
        <span>{t('ribbonListNumStyle')}</span>
        <NumFmtDropdown value={level.numFmt} onPick={(numFmt) => patch({ numFmt })} />
      </label>
      <label className="fld list-dialog-num">
        <span>{t('ribbonListStartAt')}</span>
        <input
          type="number"
          min={0}
          value={level.start ?? 1}
          onChange={(e) => patch({ start: Math.max(0, Math.floor(Number(e.target.value) || 0)) })}
        />
      </label>
      <label className="fld">
        <span>{t('ribbonListFormatText')}</span>
        <input
          value={level.lvlText}
          onChange={(e) => patch({ lvlText: e.target.value })}
          aria-label={t('ribbonListFormatText')}
        />
      </label>
      <MarkerFontFields level={level} onChange={patch} />
      <label className="fld">
        <span>{t('ribbonListAlignment')}</span>
        <Dropdown
          value={level.lvlJc ?? 'left'}
          ariaLabel={t('ribbonListAlignment')}
          options={JC_OPTIONS.map((o) => ({ value: o.value, label: t(o.labelKey) }))}
          onPick={(lvlJc) => patch({ lvlJc })}
        />
      </label>
      <LevelPreview levels={[level]} />
    </Modal>
  )
}

/* ================= Define New Multilevel List ================= */

export function DefineMultilevelDialog({
  initial,
  styles,
  onApply,
  onClose,
}: {
  initial?: CustomNumberingLevel[]
  /** paragraph styles a level can be linked to */
  styles: StyleInfo[]
  onApply: (levels: CustomNumberingLevel[]) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [levels, setLevels] = useState<CustomNumberingLevel[]>(
    initial && initial.length > 0 ? padLevels(initial) : MULTILEVEL_LIBRARY[1],
  )
  const [active, setActive] = useState(0)
  const [legal, setLegal] = useState(levels.some((l) => l.isLgl))
  const update = (i: number, patch: Partial<CustomNumberingLevel>) =>
    setLevels((ls) => ls.map((l, k) => (k === i ? { ...l, ...patch } : l)))
  const level = levels[active]
  const styleOptions = useMemo(
    () => [
      { value: '', label: t('ribbonNoStyle') },
      ...styles
        .filter((s) => s.type === 'paragraph' && !s.semiHidden)
        .map((s) => ({ value: s.styleId, label: s.name || s.styleId })),
    ],
    [styles, t],
  )
  const restartOptions = [
    { value: '-1', label: t('ribbonAuto') },
    { value: '0', label: t('ribbonRestartNever') },
    ...Array.from({ length: active }, (_, i) => ({ value: String(i + 1), label: String(i + 1) })),
  ]
  const apply = () =>
    onApply(levels.map((l) => (legal ? { ...l, isLgl: true } : { ...l, isLgl: undefined })))
  return (
    <Modal
      title={t('ribbonDefineNewList')}
      className="list-define-modal"
      onClose={onClose}
      onOk={apply}
    >
      <div className="list-define-layout">
        <div className="list-define-levels" role="listbox" aria-label={t('ribbonListLevel')}>
          {levels.map((l, i) => (
            <button
              key={i}
              type="button"
              role="option"
              aria-selected={active === i}
              className={active === i ? 'selected' : undefined}
              style={{ paddingLeft: 8 + i * 10 }}
              onClick={() => setActive(i)}
            >
              <span className="list-define-level-no">{i + 1}</span>
              <span className="list-define-level-text">
                {previewLevelText(
                  levels.map((x) => (legal ? { ...x, isLgl: true } : x)),
                  i,
                )}
              </span>
            </button>
          ))}
        </div>
        <div className="list-define-fields">
          <label className="fld">
            <span>{t('ribbonListNumStyle')}</span>
            <NumFmtDropdown
              value={level.numFmt}
              withBullet
              onPick={(numFmt) =>
                update(active, {
                  numFmt,
                  lvlText:
                    numFmt === 'bullet'
                      ? '•'
                      : level.lvlText.includes('%')
                        ? level.lvlText
                        : `%${active + 1}.`,
                })
              }
            />
          </label>
          <label className="fld">
            <span>{t('ribbonListFormatText')}</span>
            <input
              value={level.lvlText}
              aria-label={t('ribbonListFormatText')}
              onChange={(e) => update(active, { lvlText: e.target.value })}
            />
          </label>
          <div className="list-define-grid">
            <label className="fld list-dialog-num">
              <span>{t('ribbonListStartAt')}</span>
              <input
                type="number"
                min={0}
                value={level.start ?? 1}
                onChange={(e) =>
                  update(active, { start: Math.max(0, Math.floor(Number(e.target.value) || 0)) })
                }
              />
            </label>
            <label className="fld list-dialog-num">
              <span>{t('ribbonNumberPosition')}</span>
              <LengthInput
                key={`marker:${active}`}
                value={Math.max(0, level.indentLeft - (level.hanging ?? 0))}
                min={0}
                max={MAX_LIST_TWIPS}
                ariaLabel={t('ribbonNumberPosition')}
                live
                onCommit={(twips) => {
                  const marker = Math.min(twips ?? 0, level.indentLeft)
                  update(active, { hanging: level.indentLeft - marker })
                }}
              />
            </label>
            <label className="fld list-dialog-num">
              <span>{t('ribbonTextIndent')}</span>
              <LengthInput
                key={`text:${active}`}
                value={level.indentLeft}
                min={0}
                max={MAX_LIST_TWIPS}
                ariaLabel={t('ribbonTextIndent')}
                live
                onCommit={(twips) => {
                  const marker = Math.max(0, level.indentLeft - (level.hanging ?? 0))
                  const indentLeft = twips ?? 0
                  update(active, { indentLeft, hanging: Math.max(0, indentLeft - marker) })
                }}
              />
            </label>
            <label className="fld">
              <span>{t('ribbonFollowNumberWith')}</span>
              <Dropdown
                value={level.suff ?? 'tab'}
                ariaLabel={t('ribbonFollowNumberWith')}
                options={SUFF_OPTIONS.map((o) => ({ value: o.value, label: t(o.labelKey) }))}
                onPick={(suff) => update(active, { suff })}
              />
            </label>
            <label className="fld">
              <span>{t('ribbonListAlignment')}</span>
              <Dropdown
                value={level.lvlJc ?? 'left'}
                ariaLabel={t('ribbonListAlignment')}
                options={JC_OPTIONS.map((o) => ({ value: o.value, label: t(o.labelKey) }))}
                onPick={(lvlJc) => update(active, { lvlJc })}
              />
            </label>
            <label className="fld">
              <span>{t('ribbonLinkLevelToStyle')}</span>
              <Dropdown
                value={level.pStyle ?? ''}
                ariaLabel={t('ribbonLinkLevelToStyle')}
                options={styleOptions}
                onPick={(pStyle) => update(active, { pStyle: pStyle || undefined })}
              />
            </label>
            <label className="fld">
              <span>{t('ribbonRestartAfterLevel')}</span>
              <Dropdown
                value={String(level.lvlRestart ?? -1)}
                ariaLabel={t('ribbonRestartAfterLevel')}
                options={restartOptions}
                onPick={(v) => update(active, { lvlRestart: v === '-1' ? undefined : Number(v) })}
              />
            </label>
          </div>
          <MarkerFontFields level={level} onChange={(p) => update(active, p)} />
          <label className="font-check">
            <input type="checkbox" checked={legal} onChange={(e) => setLegal(e.target.checked)} />
            <span>{t('ribbonLegalStyleNumbering')}</span>
          </label>
        </div>
      </div>
    </Modal>
  )
}

function padLevels(levels: CustomNumberingLevel[]): CustomNumberingLevel[] {
  const out = [...levels]
  while (out.length < 9)
    out.push({
      numFmt: 'decimal',
      lvlText: `%${out.length + 1}.`,
      indentLeft: 720 * (out.length + 1),
      hanging: 360,
    })
  return out.slice(0, 9)
}
