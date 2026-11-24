import { useState } from 'react'
import {
  REL_SIZE_FROM_H,
  REL_SIZE_FROM_V,
  relSizeBaseTwips,
  type RelSize,
  type RelSizeFrom,
  type SectionSettings,
} from '@genoffice/docx-engine'
import type { ShapeSide, ShapeSizeGeometry } from '../editor/shape-size'
import { useI18n, type StringKey } from '../i18n/locale'
import { pxToTwips, twipsToPx } from '../units'
import { LengthInput } from './LengthInput'
import { FormDialog } from './MenuDialogs'

const SIDE_TWIPS_MIN = 14
const SIDE_TWIPS_MAX = 31680
const PCT_MIN = 1
const PCT_MAX = 1000

const FROM_LABELS: Record<RelSizeFrom, StringKey> = {
  margin: 'ribbonTcBalloonMargin',
  page: 'ribbonDocPage',
  leftMargin: 'appRulerLeftMargin',
  rightMargin: 'appRulerRightMargin',
  topMargin: 'appRulerTopMargin',
  bottomMargin: 'appRulerBottomMargin',
  insideMargin: 'ribbonMarginInside',
  outsideMargin: 'ribbonMarginOutside',
}

interface SideState {
  px: number
  relative: boolean
  /** the share as typed: it stays as typed until it is a number the side can take */
  pct: string
  from: RelSizeFrom
}

const pctText = (pct: number) => String(Math.round(pct * 10) / 10)

function sideState(
  side: ShapeSide,
  axis: 'H' | 'V',
  section: SectionSettings | null | undefined,
): SideState {
  const from = side.rel?.from ?? 'margin'
  const base = section ? relSizeBaseTwips(section, axis, from) : 0
  return {
    px: side.px,
    relative: !!side.rel,
    pct: pctText(side.rel?.pct ?? (base > 0 ? (pxToTwips(side.px) / base) * 100 : 100)),
    from,
  }
}

function relOf(state: SideState): RelSize | null {
  const pct = Number(state.pct)
  if (state.pct.trim() === '' || !Number.isFinite(pct) || pct < PCT_MIN || pct > PCT_MAX)
    return null
  return { from: state.from, pct }
}

/** Word's Size tab for a shape or a text box: each side absolute, or a share of the page */
export function ShapeSizeDialog({
  current,
  section,
  title,
  onApply,
  onClose,
}: {
  current: ShapeSizeGeometry
  /** the section the shape stands in: what a share is a share of */
  section: SectionSettings | null | undefined
  title: string
  onApply: (next: ShapeSizeGeometry) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [w, setW] = useState(() => sideState(current.w, 'H', section))
  const [h, setH] = useState(() => (current.h ? sideState(current.h, 'V', section) : null))
  const valid = (state: SideState | null) => !state || !state.relative || relOf(state) !== null
  const result = (state: SideState): ShapeSide => ({
    px: state.px,
    rel: state.relative ? relOf(state) : null,
  })
  const side = (
    axis: 'H' | 'V',
    state: SideState,
    set: (next: SideState) => void,
    label: string,
  ) => {
    const key = axis === 'H' ? 'width' : 'height'
    return (
      <fieldset data-side={key}>
        <legend>{label}</legend>
        <div className="tco-grid shape-size-grid">
          <label className="tco-check">
            <input
              type="radio"
              name={`${key}Mode`}
              value="absolute"
              checked={!state.relative}
              onChange={() => set({ ...state, relative: false })}
            />
            {t('ribbonSizeAbsolute')}
          </label>
          <LengthInput
            value={pxToTwips(state.px)}
            min={SIDE_TWIPS_MIN}
            max={SIDE_TWIPS_MAX}
            disabled={state.relative}
            ariaLabel={`${label}: ${t('ribbonSizeAbsolute')}`}
            live
            onCommit={(twips) => twips && set({ ...state, px: twipsToPx(twips) })}
          />
          <label className="tco-check">
            <input
              type="radio"
              name={`${key}Mode`}
              value="relative"
              disabled={!section}
              checked={state.relative}
              onChange={() => set({ ...state, relative: true })}
            />
            {t('ribbonSizeRelative')}
          </label>
          <span className="size-scale-field">
            <input
              name={`${key}Pct`}
              inputMode="decimal"
              disabled={!state.relative}
              aria-label={`${label}: ${t('ribbonSizeRelative')}`}
              value={state.pct}
              onChange={(e) => set({ ...state, pct: e.target.value })}
            />
            %
          </span>
          <span>{t('ribbonSizeRelativeTo')}</span>
          <select
            name={`${key}From`}
            disabled={!state.relative}
            aria-label={`${label}: ${t('ribbonSizeRelativeTo')}`}
            value={state.from}
            onChange={(e) => set({ ...state, from: e.target.value as RelSizeFrom })}
          >
            {(axis === 'H' ? REL_SIZE_FROM_H : REL_SIZE_FROM_V).map((from) => (
              <option key={from} value={from}>
                {t(FROM_LABELS[from])}
              </option>
            ))}
          </select>
        </div>
      </fieldset>
    )
  }
  return (
    <FormDialog
      title={title}
      name="size-position-dialog shape-size-dialog"
      okDisabled={!valid(w) || !valid(h)}
      onClose={onClose}
      onSubmit={() => {
        onClose()
        onApply({ w: result(w), h: h ? result(h) : null })
      }}
    >
      {h && side('V', h, setH, t('ribbonPicHeight'))}
      {side('H', w, setW, t('ribbonPicWidth'))}
    </FormDialog>
  )
}
