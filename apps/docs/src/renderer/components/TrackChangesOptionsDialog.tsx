import { useState } from 'react'
import { useI18n } from '../i18n/locale'
import {
  BALLOON_WIDTH_MAX,
  BALLOON_WIDTH_MIN,
  CELL_COLOR_NAMES,
  DEFAULT_TRACK_CHANGES_OPTIONS,
  REV_COLOR_NAMES,
  clampBalloonWidth,
  type CellColorName,
  type ChangedLines,
  type LineColor,
  type RevColor,
  type RevMark,
  type TrackChangesOptions,
} from '../editor/track-changes-options'
import { useModalKeys } from './modal-keys'

type TKey = Parameters<ReturnType<typeof useI18n>['t']>[0]

const MARK_KEYS: Record<RevMark, TKey> = {
  none: 'ribbonTcMarkNone',
  color: 'ribbonTcMarkColorOnly',
  bold: 'ribbonTcMarkBold',
  italic: 'ribbonTcMarkItalic',
  underline: 'ribbonTcMarkUnderline',
  dunderline: 'ribbonTcMarkDoubleUnderline',
  strike: 'ribbonTcMarkStrike',
  dstrike: 'ribbonTcMarkDoubleStrike',
}

const COLOR_KEYS: Record<CellColorName, TKey> = {
  black: 'ribbonTcColorBlack',
  blue: 'ribbonTcColorBlue',
  turquoise: 'ribbonTcColorTurquoise',
  brightGreen: 'ribbonTcColorBrightGreen',
  pink: 'ribbonTcColorPink',
  red: 'ribbonTcColorRed',
  yellow: 'ribbonTcColorYellow',
  white: 'ribbonTcColorWhite',
  darkBlue: 'ribbonTcColorDarkBlue',
  teal: 'ribbonTcColorTeal',
  green: 'ribbonTcColorGreen',
  violet: 'ribbonTcColorViolet',
  darkRed: 'ribbonTcColorDarkRed',
  darkYellow: 'ribbonTcColorDarkYellow',
  gray50: 'ribbonTcColorGray50',
  gray25: 'ribbonTcColorGray25',
  classicRed: 'ribbonTcColorClassicRed',
  classicBlue: 'ribbonTcColorClassicBlue',
  lightBlue: 'ribbonTcColorLightBlue',
  palePink: 'ribbonTcColorPalePink',
  lightYellow: 'ribbonTcColorLightYellow',
  lightOrange: 'ribbonTcColorLightOrange',
}

const LINE_KEYS: Record<ChangedLines, TKey> = {
  outside: 'ribbonTcLinesOutside',
  left: 'ribbonTcLinesLeft',
  right: 'ribbonTcLinesRight',
  none: 'ribbonTcMarkNone',
}

const MARKS = Object.keys(MARK_KEYS) as RevMark[]

interface Props {
  options: TrackChangesOptions
  onApply: (options: TrackChangesOptions) => void
  onClose: () => void
}

/** Word's Review ▸ Tracking ▸ Track Changes Options ▸ Advanced Options. */
export function TrackChangesOptionsDialog({ options, onApply, onClose }: Props) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [o, setO] = useState<TrackChangesOptions>(options)
  const [width, setWidth] = useState(String(options.balloonWidth))
  const set = <K extends keyof TrackChangesOptions>(k: K, v: TrackChangesOptions[K]) =>
    setO((prev) => ({ ...prev, [k]: v }))

  const markSelect = (k: keyof TrackChangesOptions, name: string) => (
    <select name={name} value={o[k] as RevMark} onChange={(e) => set(k, e.target.value as never)}>
      {MARKS.map((m) => (
        <option key={m} value={m}>
          {t(MARK_KEYS[m])}
        </option>
      ))}
    </select>
  )
  const colorSelect = (
    k: keyof TrackChangesOptions,
    name: string,
    first: { value: RevColor | LineColor; label: TKey } | null,
    names: readonly CellColorName[] = REV_COLOR_NAMES,
  ) => (
    <select name={name} value={o[k] as string} onChange={(e) => set(k, e.target.value as never)}>
      {first && <option value={first.value}>{t(first.label)}</option>}
      {names.map((c) => (
        <option key={c} value={c}>
          {t(COLOR_KEYS[c])}
        </option>
      ))}
    </select>
  )
  const byAuthor = { value: 'author', label: 'ribbonTcColorByAuthor' } as const
  const colorLabel = <span className="tco-color-cell">{t('ribbonTcColor')}</span>
  const row = (
    label: TKey,
    markK: keyof TrackChangesOptions,
    colorK: keyof TrackChangesOptions,
    name: string,
  ) => (
    <>
      <span>{t(label)}</span>
      {markSelect(markK, `${name}-mark`)}
      {colorLabel}
      {colorSelect(colorK, `${name}-color`, byAuthor)}
    </>
  )
  const check = (label: TKey, k: 'trackMoves' | 'trackFormatting' | 'balloonLines') => (
    <label className="tco-check">
      <input type="checkbox" checked={o[k]} onChange={(e) => set(k, e.target.checked)} />
      {t(label)}
    </label>
  )
  const submit = () => {
    onApply({ ...o, balloonWidth: clampBalloonWidth(Number(width)) })
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          submit()
        }}
      >
        <h2>{t('ribbonTcOptionsTitle')}</h2>
        <fieldset>
          <legend>{t('ribbonTcMarkup')}</legend>
          <div className="tco-grid">
            {row('ribbonTcInsertions', 'insMark', 'insColor', 'ins')}
            {row('ribbonTcDeletions', 'delMark', 'delColor', 'del')}
            <span>{t('ribbonTcChangedLines')}</span>
            <select
              name="lines-position"
              value={o.changedLines}
              onChange={(e) => set('changedLines', e.target.value as ChangedLines)}
            >
              {(Object.keys(LINE_KEYS) as ChangedLines[]).map((v) => (
                <option key={v} value={v}>
                  {t(LINE_KEYS[v])}
                </option>
              ))}
            </select>
            {colorLabel}
            {colorSelect('changedLinesColor', 'lines-color', {
              value: 'auto',
              label: 'ribbonAuto',
            })}
            <span>{t('ribbonTcComments')}</span>
            <span />
            {colorLabel}
            {colorSelect('commentColor', 'comment-color', byAuthor)}
          </div>
        </fieldset>
        <fieldset>
          <legend>{t('ribbonTcMoves')}</legend>
          {check('ribbonTcTrackMoves', 'trackMoves')}
          <div className="tco-grid">
            {row('ribbonTcMovedFrom', 'moveFromMark', 'moveFromColor', 'move-from')}
            {row('ribbonTcMovedTo', 'moveToMark', 'moveToColor', 'move-to')}
          </div>
        </fieldset>
        <fieldset>
          <legend>{t('ribbonTcCellHighlight')}</legend>
          <div className="tco-grid">
            <span>{t('ribbonTcInsertedCells')}</span>
            {colorSelect('cellIns', 'cell-ins', null, CELL_COLOR_NAMES)}
            <span>{t('ribbonTcMergedCells')}</span>
            {colorSelect('cellMerge', 'cell-merge', null, CELL_COLOR_NAMES)}
            <span>{t('ribbonTcDeletedCells')}</span>
            {colorSelect('cellDel', 'cell-del', null, CELL_COLOR_NAMES)}
            <span>{t('ribbonTcSplitCells')}</span>
            {colorSelect('cellSplit', 'cell-split', null, CELL_COLOR_NAMES)}
          </div>
        </fieldset>
        <fieldset>
          <legend>{t('ribbonTcFormatting')}</legend>
          {check('ribbonTcTrackFormatting', 'trackFormatting')}
          <div className="tco-grid">{row('ribbonTcFormatting', 'fmtMark', 'fmtColor', 'fmt')}</div>
        </fieldset>
        <fieldset>
          <legend>{t('ribbonTcBalloons')}</legend>
          <div className="tco-grid">
            <span>{t('ribbonTcBalloonWidth')}</span>
            <input
              type="number"
              name="balloon-width"
              min={BALLOON_WIDTH_MIN}
              max={BALLOON_WIDTH_MAX}
              step={1}
              value={width}
              onChange={(e) => setWidth(e.target.value)}
              onBlur={() => setWidth(String(clampBalloonWidth(Number(width))))}
            />
            <span className="tco-color-cell">{t('ribbonTcBalloonMargin')}</span>
            <select
              name="balloon-margin"
              value={o.balloonMargin}
              onChange={(e) => set('balloonMargin', e.target.value as 'right' | 'left')}
            >
              <option value="right">{t('ribbonTcMarginRight')}</option>
              <option value="left">{t('ribbonTcMarginLeft')}</option>
            </select>
          </div>
          {check('ribbonTcBalloonLines', 'balloonLines')}
        </fieldset>
        <div className="modal-actions">
          <button
            type="button"
            onClick={() => {
              setO({ ...DEFAULT_TRACK_CHANGES_OPTIONS })
              setWidth(String(DEFAULT_TRACK_CHANGES_OPTIONS.balloonWidth))
            }}
          >
            {t('ribbonTcReset')}
          </button>
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
