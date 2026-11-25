import { useState } from 'react'
import type { DrawingGrid } from '@genoffice/docx-engine'
import {
  DEFAULT_GRID_DISPLAY_EVERY,
  DEFAULT_GRID_ORIGIN_TWIPS,
  DEFAULT_GRID_SPACING_TWIPS,
} from '../editor/drawing-grid'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

export interface DrawingGridResult {
  /** the grid to write, when a grid field was changed */
  grid?: DrawingGrid
  snap: boolean
  show: boolean
}

interface Fields {
  hSpacing: number
  vSpacing: number
  useMargins: boolean
  hOrigin: number
  vOrigin: number
  verticalLines: boolean
  vEvery: number
  hEvery: number
}

const MIN_SPACING_PT = 1
const MAX_SPACING_PT = 1584
const MAX_EVERY = 127

const fieldsOf = (grid: DrawingGrid | null): Fields => {
  const spacing = (twips: number | undefined) =>
    (twips && twips >= MIN_SPACING_PT * 20 ? twips : DEFAULT_GRID_SPACING_TWIPS) / 20
  const vertical = grid?.displayHorizontalEvery ?? DEFAULT_GRID_DISPLAY_EVERY
  return {
    hSpacing: spacing(grid?.horizontalSpacing),
    vSpacing: spacing(grid?.verticalSpacing),
    useMargins: !grid?.ownOrigin,
    hOrigin: (grid?.horizontalOrigin ?? DEFAULT_GRID_ORIGIN_TWIPS.horizontal) / 20,
    vOrigin: (grid?.verticalOrigin ?? DEFAULT_GRID_ORIGIN_TWIPS.vertical) / 20,
    verticalLines: vertical > 0,
    vEvery: Math.max(1, vertical),
    hEvery: Math.max(1, grid?.displayVerticalEvery ?? DEFAULT_GRID_DISPLAY_EVERY),
  }
}

/** the file's grid with the fields the dialog changed; undefined when none was */
export function drawingGridChanges(
  grid: DrawingGrid | null,
  before: Fields,
  after: Fields,
): DrawingGrid | undefined {
  const out: DrawingGrid = { ...grid }
  let changed = false
  const set = <K extends keyof DrawingGrid>(key: K, value: DrawingGrid[K], dirty: boolean) => {
    if (!dirty) return
    changed = true
    if (value === undefined) delete out[key]
    else out[key] = value
  }
  set('horizontalSpacing', Math.round(after.hSpacing * 20), after.hSpacing !== before.hSpacing)
  set('verticalSpacing', Math.round(after.vSpacing * 20), after.vSpacing !== before.vSpacing)
  set(
    'displayHorizontalEvery',
    after.verticalLines ? after.vEvery : 0,
    after.verticalLines !== before.verticalLines || after.vEvery !== before.vEvery,
  )
  set('displayVerticalEvery', after.hEvery, after.hEvery !== before.hEvery)
  set('ownOrigin', after.useMargins ? undefined : true, after.useMargins !== before.useMargins)
  const originDirty =
    !after.useMargins &&
    (before.useMargins || after.hOrigin !== before.hOrigin || after.vOrigin !== before.vOrigin)
  set('horizontalOrigin', Math.round(after.hOrigin * 20), originDirty)
  set('verticalOrigin', Math.round(after.vOrigin * 20), originDirty)
  return changed ? out : undefined
}

/** Word's Format > Document > Document Grid > Drawing Grid */
export function DrawingGridDialog({
  grid,
  snap,
  show,
  onApply,
  onClose,
}: {
  grid: DrawingGrid | null
  snap: boolean
  show: boolean
  onApply: (result: DrawingGridResult) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [opened] = useState(() => fieldsOf(grid))
  const [value, setValue] = useState(opened)
  const [snapOn, setSnapOn] = useState(snap)
  const [showOn, setShowOn] = useState(show)
  const valid =
    [value.hSpacing, value.vSpacing].every((v) => v >= MIN_SPACING_PT && v <= MAX_SPACING_PT) &&
    [value.vEvery, value.hEvery].every((v) => Number.isInteger(v) && v >= 1 && v <= MAX_EVERY) &&
    [value.hOrigin, value.vOrigin].every((v) => v >= 0 && v <= MAX_SPACING_PT)

  const length = (key: 'hSpacing' | 'vSpacing' | 'hOrigin' | 'vOrigin', label: string) => (
    <>
      <span>{label}</span>
      <span className="doc-dialog-pitch">
        <input
          type="number"
          name={key}
          aria-label={label}
          min={key.endsWith('Spacing') ? MIN_SPACING_PT : 0}
          max={MAX_SPACING_PT}
          step={0.05}
          disabled={key.endsWith('Origin') && value.useMargins}
          value={Number.isNaN(value[key]) ? '' : value[key]}
          onChange={(e) => setValue((v) => ({ ...v, [key]: e.target.valueAsNumber }))}
        />
        <span className="para-unit">{t('appUnitPtAbbr')}</span>
      </span>
    </>
  )
  const every = (key: 'vEvery' | 'hEvery', label: string, disabled = false) => (
    <input
      type="number"
      name={key}
      aria-label={label}
      min={1}
      max={MAX_EVERY}
      step={1}
      disabled={disabled}
      value={Number.isNaN(value[key]) ? '' : value[key]}
      onChange={(e) => setValue((v) => ({ ...v, [key]: e.target.valueAsNumber }))}
    />
  )

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog draw-grid-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          e.stopPropagation()
          if (!valid) return
          onApply({
            grid: drawingGridChanges(grid, opened, value),
            snap: snapOn,
            show: showOn,
          })
          onClose()
        }}
      >
        <h2>{t('ribbonDocGridDrawing').replace(/\s*(?:…|\.\.\.)$/, '')}</h2>
        <fieldset>
          <legend>{t('ribbonDrawGridSnap')}</legend>
          <label className="tco-check">
            <input
              type="checkbox"
              name="snapToGrid"
              checked={snapOn}
              onChange={(e) => setSnapOn(e.target.checked)}
            />
            {t('ribbonDrawGridSnapToGrid')}
          </label>
        </fieldset>
        <fieldset>
          <legend>{t('ribbonDrawGridSettings')}</legend>
          <div className="tco-grid draw-grid-fields">
            {length('hSpacing', t('ribbonDrawGridHSpacing'))}
            {length('vSpacing', t('ribbonDrawGridVSpacing'))}
          </div>
        </fieldset>
        <fieldset>
          <legend>{t('ribbonDrawGridOrigin')}</legend>
          <label className="tco-check">
            <input
              type="checkbox"
              name="useMargins"
              checked={value.useMargins}
              onChange={(e) => setValue((v) => ({ ...v, useMargins: e.target.checked }))}
            />
            {t('ribbonDrawGridUseMargins')}
          </label>
          <div className="tco-grid draw-grid-fields">
            {length('hOrigin', t('ribbonDrawGridHOrigin'))}
            {length('vOrigin', t('ribbonDrawGridVOrigin'))}
          </div>
        </fieldset>
        <fieldset>
          <legend>{t('ribbonDrawGridShow')}</legend>
          <label className="tco-check">
            <input
              type="checkbox"
              name="displayGridlines"
              checked={showOn}
              onChange={(e) => setShowOn(e.target.checked)}
            />
            {t('ribbonDrawGridDisplay')}
          </label>
          <div className="tco-grid draw-grid-fields">
            <label className="tco-check">
              <input
                type="checkbox"
                name="verticalLines"
                checked={value.verticalLines}
                onChange={(e) => setValue((v) => ({ ...v, verticalLines: e.target.checked }))}
              />
              {t('ribbonDrawGridVEvery')}
            </label>
            {every('vEvery', t('ribbonDrawGridVEvery'), !value.verticalLines)}
            <span>{t('ribbonDrawGridHEvery')}</span>
            {every('hEvery', t('ribbonDrawGridHEvery'))}
          </div>
        </fieldset>
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="submit" className="primary" disabled={!valid}>
            {t('ribbonOk')}
          </button>
        </div>
      </form>
    </div>
  )
}

export const drawingGridFieldsOf = fieldsOf
