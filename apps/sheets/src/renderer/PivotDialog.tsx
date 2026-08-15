import { useState } from 'react'

import { useI18n, type StringKey } from './i18n/locale'

/// Native OOXML pivot table dialog: creates a real PivotTable part (not
/// formula-based) via the same add_pivot pathway the AI assistant uses.
/// The caller supplies the current source-range headers so the user can
/// pick which field is the row dimension, which is the column dimension,
/// and which value(s) to aggregate.

export interface PivotField {
  readonly label: string
  readonly colIndex: number
}

/// Legacy v1 interface kept for SubtotalDialog which reuses PivotField.
export interface PivotConfig {
  readonly categoryCol: number
  readonly valueCol: number
  readonly agg: 'sum' | 'count' | 'average'
}

export type PivotShowDataAsOption = 'percentOfTotal' | 'percentOfRow' | 'percentOfCol'

/// Grouping mode for dimension fields: dates by year/quarter/month, numbers by
/// fixed-step intervals.
export type PivotGroupingOption =
  | { readonly kind: 'date'; readonly dateUnit: 'year' | 'quarter' | 'month' }
  | { readonly kind: 'range'; readonly rangeStep: number }

/// Label filter for row/column fields (equals/contains/begins-with).
export interface PivotLabelFilterOption {
  readonly op: 'equal' | 'contains' | 'beginsWith'
  readonly value: string
}

/// Value filter on a data field (top N / greater than / between), applied to
/// level-1 row-field members.
export interface PivotValueFilterOption {
  readonly op: 'top' | 'greaterThan' | 'between'
  readonly count?: number | undefined
  readonly from?: number | undefined
  readonly to?: number | undefined
}

export interface PivotValueSpec {
  readonly fieldIndex: number // index into PivotField[]; -1 for calculated fields
  readonly agg: 'sum' | 'count' | 'average' | 'max' | 'min'
  /// "Show values as" mode; undefined = normal (show the aggregate directly).
  readonly showDataAs?: PivotShowDataAsOption | undefined
  /// Calculated field: name + formula (basic arithmetic over other field names,
  /// e.g. Revenue-Cost).
  readonly calcName?: string | undefined
  readonly formula?: string | undefined
}

/// Dialog prefill context assembled by App when editing an existing pivot (A3).
export interface PivotEditSeed {
  readonly fields: readonly PivotField[]
  readonly sourceRange: string
  readonly initial: {
    readonly rowFieldIndices: readonly number[]
    readonly colFieldIndices: readonly number[]
    readonly values: readonly PivotValueSpec[]
    readonly targetCell: string
  }
}

/// Full OOXML pivot configuration passed to onCreate on success.
export interface OoXmlPivotConfig {
  readonly sourceRange: string // e.g. "A1:C10"
  /// Row dimension fields (ordered, outer first), indices into fields; at least 1.
  readonly rowFieldIndices: readonly number[]
  /// Column dimension fields (ordered, outer first), indices into fields; empty
  /// array = no column dimension.
  readonly colFieldIndices: readonly number[]
  /// Grouping modes of dimension fields (fieldIndex points into fields; only
  /// row/column dimension fields).
  readonly groupings: readonly { readonly fieldIndex: number; readonly rule: PivotGroupingOption }[]
  /// Label filters on row/column fields (fieldIndex points into fields).
  readonly labelFilters: readonly {
    readonly fieldIndex: number
    readonly rule: PivotLabelFilterOption
  }[]
  /// Value filter (valueIndex points into values; applied to the level-1 row
  /// field).
  readonly valueFilters: readonly {
    readonly valueIndex: number
    readonly rule: PivotValueFilterOption
  }[]
  readonly values: readonly PivotValueSpec[]
  readonly targetCell: string // e.g. "F1"
}

const AGG_LABELS: Record<PivotValueSpec['agg'], StringKey> = {
  sum: 'dlgPivotAggSum',
  count: 'dlgPivotAggCount',
  average: 'dlgPivotAggAverage',
  max: 'dlgPivotAggMax',
  min: 'dlgPivotAggMin',
}

/// "Show values as" options; '' means normal (no showDataAs set).
const SHOW_DATA_AS_LABELS: readonly { value: PivotShowDataAsOption | ''; labelKey: StringKey }[] = [
  { value: '', labelKey: 'dlgPivotShowAsNormal' },
  { value: 'percentOfTotal', labelKey: 'dlgPivotShowAsPctTotal' },
  { value: 'percentOfRow', labelKey: 'dlgPivotShowAsPctRow' },
  { value: 'percentOfCol', labelKey: 'dlgPivotShowAsPctCol' },
]

export function PivotDialog({
  fields,
  sourceRange,
  mode = 'create',
  initial,
  onCreate,
  onClose,
}: {
  readonly fields: readonly PivotField[]
  /// Pre-filled source range from the current selection.
  readonly sourceRange: string
  /// 'edit' edits an existing pivot: layout prefilled, target cell locked to the
  /// original output area's top-left corner.
  readonly mode?: 'create' | 'edit'
  readonly initial?: PivotEditSeed['initial'] | undefined
  /// Returns an error message, or null on success.
  readonly onCreate: (config: OoXmlPivotConfig) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  // Row/column dimensions: ordered multi-select (outer first) with add/remove and
  // move up/down; column dimensions may be empty.
  const [rowFieldIndices, setRowFieldIndices] = useState<number[]>([
    ...(initial?.rowFieldIndices ?? [0]),
  ])
  const [colFieldIndices, setColFieldIndices] = useState<number[]>([
    ...(initial?.colFieldIndices ?? []),
  ])
  // Grouping modes of dimension fields, recorded by field index (independent of
  // axis/level, so moving levels doesn't lose them).
  const [fieldGroupings, setFieldGroupings] = useState<Record<number, PivotGroupingOption>>({})
  // Label filters on row/column fields (by field index) and value filters on data
  // fields (by value ordinal).
  const [labelFilters, setLabelFilters] = useState<Record<number, PivotLabelFilterOption>>({})
  const [valueFilters, setValueFilters] = useState<Record<number, PivotValueFilterOption>>({})
  const [values, setValues] = useState<PivotValueSpec[]>(
    initial?.values !== undefined
      ? [...initial.values]
      : [{ fieldIndex: fields.length > 1 ? 1 : 0, agg: 'sum' }],
  )
  const [targetCell, setTargetCell] = useState(initial?.targetCell ?? '')
  const [error, setError] = useState<string | null>(null)

  const noFields = fields.length === 0

  const updateValue = (index: number, patch: Partial<PivotValueSpec>) => {
    setValues((prev) => prev.map((v, i) => (i === index ? { ...v, ...patch } : v)))
  }
  const addValue = () => {
    setValues((prev) => [...prev, { fieldIndex: 0, agg: 'sum' }])
  }
  // Calculated-field row: name + formula inputs (aggregation fixed to sum; the
  // formula operates on in-group sums).
  const addCalculatedValue = () => {
    setValues((prev) => [...prev, { fieldIndex: -1, agg: 'sum', calcName: '', formula: '' }])
  }
  const removeValue = (index: number) => {
    setValues((prev) => prev.filter((_, i) => i !== index))
  }

  // Ordered field-list operations shared by rows/columns (the call site's setter
  // decides which axis they apply to).
  type AxisSetter = (updater: (prev: number[]) => number[]) => void
  const updateAxisField = (setAxis: AxisSetter, index: number, fieldIdx: number) => {
    setAxis((prev) => prev.map((v, i) => (i === index ? fieldIdx : v)))
  }
  const addAxisField = (setAxis: AxisSetter) => {
    // Default to the first field not yet selected on this axis.
    setAxis((prev) => {
      const unused = fields.findIndex((_, i) => !prev.includes(i))
      return [...prev, unused >= 0 ? unused : 0]
    })
  }
  const removeAxisField = (setAxis: AxisSetter, index: number) => {
    setAxis((prev) => prev.filter((_, i) => i !== index))
  }
  const moveAxisField = (setAxis: AxisSetter, index: number, delta: -1 | 1) => {
    setAxis((prev) => {
      const target = index + delta
      if (target < 0 || target >= prev.length) return prev
      const next = [...prev]
      const tmp = next[index]!
      next[index] = next[target]!
      next[target] = tmp
      return next
    })
  }

  // Grouping dropdown encoded value ↔ rule (range step is maintained separately
  // via a numeric input).
  const groupingSelectValue = (rule: PivotGroupingOption | undefined): string => {
    if (!rule) return ''
    return rule.kind === 'date' ? `date:${rule.dateUnit}` : 'range'
  }
  const setFieldGrouping = (fieldIdx: number, selected: string) => {
    setFieldGroupings((prev) => {
      const next = { ...prev }
      if (selected === '') {
        delete next[fieldIdx]
      } else if (selected === 'range') {
        next[fieldIdx] = { kind: 'range', rangeStep: 100 }
      } else {
        next[fieldIdx] = {
          kind: 'date',
          dateUnit: selected.slice('date:'.length) as 'year' | 'quarter' | 'month',
        }
      }
      return next
    })
  }

  /// Ordered field-list UI for one dimension axis (shared by rows/columns): one
  /// row per level, supporting field change, grouping mode, move up/down, and
  /// delete; the delete button is hidden below minCount (rows keep ≥1 level).
  const renderAxisFields = (indices: readonly number[], setAxis: AxisSetter, minCount: number) => (
    <>
      {indices.map((fieldIdx, i) => (
        <div
          key={i}
          style={{
            display: 'flex',
            flexWrap: 'wrap',
            gap: 6,
            marginBottom: 4,
            alignItems: 'center',
          }}
        >
          <span style={{ fontSize: 12, color: '#888', width: 34 }}>
            {t('dlgPivotLevel', { n: i + 1 })}
          </span>
          <select
            value={fieldIdx}
            onChange={(event) => updateAxisField(setAxis, i, Number(event.target.value))}
            style={{ flex: 1, minWidth: 0 }}
          >
            {fields.map((field, fi) => (
              <option key={field.colIndex} value={fi}>
                {field.label}
              </option>
            ))}
          </select>
          <select
            data-tip={t('dlgPivotGrouping')}
            value={groupingSelectValue(fieldGroupings[fieldIdx])}
            onChange={(event) => setFieldGrouping(fieldIdx, event.target.value)}
          >
            <option value="">{t('dlgPivotGroupNone')}</option>
            <option value="date:year">{t('dlgPivotGroupYear')}</option>
            <option value="date:quarter">{t('dlgPivotGroupQuarter')}</option>
            <option value="date:month">{t('dlgPivotGroupMonth')}</option>
            <option value="range">{t('dlgPivotGroupRange')}</option>
          </select>
          {fieldGroupings[fieldIdx]?.kind === 'range' && (
            <input
              type="number"
              data-tip={t('dlgPivotRangeStep')}
              className="cell-input"
              style={{ width: 64 }}
              min={1}
              value={(fieldGroupings[fieldIdx] as { rangeStep: number }).rangeStep}
              onChange={(event) => {
                const step = Number(event.target.value)
                setFieldGroupings((prev) => ({
                  ...prev,
                  [fieldIdx]: { kind: 'range', rangeStep: step },
                }))
              }}
            />
          )}
          <select
            data-tip={t('dlgPivotLabelFilter')}
            value={labelFilters[fieldIdx]?.op ?? ''}
            onChange={(event) => {
              const op = event.target.value as PivotLabelFilterOption['op'] | ''
              setLabelFilters((prev) => {
                const next = { ...prev }
                if (op === '') delete next[fieldIdx]
                else next[fieldIdx] = { op, value: prev[fieldIdx]?.value ?? '' }
                return next
              })
            }}
          >
            <option value="">{t('dlgPivotFilterNone')}</option>
            <option value="equal">{t('dlgPivotFilterEqual')}</option>
            <option value="contains">{t('dlgPivotFilterContains')}</option>
            <option value="beginsWith">{t('dlgPivotFilterBeginsWith')}</option>
          </select>
          {labelFilters[fieldIdx] && (
            <input
              type="text"
              data-tip={t('dlgPivotLabelFilterText')}
              className="cell-input"
              placeholder={t('dlgPivotFilterTextPlaceholder')}
              style={{ width: 76 }}
              value={labelFilters[fieldIdx].value}
              onChange={(event) => {
                const value = event.target.value
                setLabelFilters((prev) => ({
                  ...prev,
                  [fieldIdx]: { op: prev[fieldIdx]!.op, value },
                }))
              }}
            />
          )}
          <button
            type="button"
            className="secondary"
            style={{ padding: '1px 6px' }}
            disabled={i === 0}
            data-tip={t('dlgPivotMoveUp')}
            aria-label={t('dlgPivotMoveUp')}
            onClick={() => moveAxisField(setAxis, i, -1)}
          >
            ↑
          </button>
          <button
            type="button"
            className="secondary"
            style={{ padding: '1px 6px' }}
            disabled={i === indices.length - 1}
            data-tip={t('dlgPivotMoveDown')}
            aria-label={t('dlgPivotMoveDown')}
            onClick={() => moveAxisField(setAxis, i, 1)}
          >
            ↓
          </button>
          {indices.length > minCount && (
            <button
              type="button"
              className="secondary"
              style={{ padding: '1px 6px' }}
              data-tip={t('dlgPivotRemoveLevel')}
              aria-label={t('dlgPivotRemoveLevel')}
              onClick={() => removeAxisField(setAxis, i)}
            >
              ✕
            </button>
          )}
        </div>
      ))}
      {indices.length < Math.min(8, fields.length) && (
        <button
          type="button"
          className="secondary"
          style={{ fontSize: 12, marginTop: 2 }}
          onClick={() => addAxisField(setAxis)}
        >
          {t('dlgPivotAddField')}
        </button>
      )}
    </>
  )

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog pivot-dialog"
        role="dialog"
        aria-label={mode === 'edit' ? t('dlgPivotEditTitle') : t('dlgPivotCreateTitle')}
        style={{ minWidth: 380 }}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{mode === 'edit' ? t('dlgPivotEditTitle') : t('dlgPivotCreateTitle')}</header>
        {noFields ? (
          <p className="dialog-note">{t('dlgPivotNoFields')}</p>
        ) : (
          <div className="dialog-grid">
            <label>
              {t('dlgPivotSourceRange')}
              <input
                type="text"
                className="cell-input"
                value={sourceRange}
                readOnly
                style={{ background: '#f5f5f5', cursor: 'default' }}
              />
            </label>
            <fieldset
              className="dialog-span"
              style={{ border: '1px solid #ddd', padding: '6px 8px', marginTop: 4 }}
            >
              <legend style={{ fontWeight: 500, fontSize: 13 }}>{t('dlgPivotRowFields')}</legend>
              {renderAxisFields(rowFieldIndices, setRowFieldIndices, 1)}
            </fieldset>
            <fieldset
              className="dialog-span"
              style={{ border: '1px solid #ddd', padding: '6px 8px', marginTop: 4 }}
            >
              <legend style={{ fontWeight: 500, fontSize: 13 }}>{t('dlgPivotColFields')}</legend>
              {colFieldIndices.length === 0 && (
                <p className="dialog-note" style={{ margin: '2px 0' }}>
                  {t('dlgPivotNoColFields')}
                </p>
              )}
              {renderAxisFields(colFieldIndices, setColFieldIndices, 0)}
            </fieldset>
            <fieldset
              className="dialog-span"
              style={{ border: '1px solid #ddd', padding: '6px 8px', marginTop: 4 }}
            >
              <legend style={{ fontWeight: 500, fontSize: 13 }}>{t('dlgPivotValues')}</legend>
              {values.map((spec, i) => (
                <div
                  key={i}
                  style={{
                    display: 'flex',
                    flexWrap: 'wrap',
                    gap: 6,
                    marginBottom: 4,
                    alignItems: 'center',
                  }}
                >
                  {spec.formula !== undefined ? (
                    // Calculated field: name + formula (e.g. Profit =
                    // Revenue-Cost).
                    <>
                      <input
                        type="text"
                        className="cell-input"
                        placeholder={t('dlgPivotCalcNamePlaceholder')}
                        value={spec.calcName ?? ''}
                        onChange={(event) => updateValue(i, { calcName: event.target.value })}
                        style={{ width: 90 }}
                      />
                      <span style={{ color: '#888' }}>=</span>
                      <input
                        type="text"
                        className="cell-input"
                        placeholder={t('dlgPivotCalcFormulaPlaceholder')}
                        value={spec.formula}
                        onChange={(event) => updateValue(i, { formula: event.target.value })}
                        style={{ flex: 1, minWidth: 0 }}
                      />
                    </>
                  ) : (
                    <>
                      <select
                        value={spec.fieldIndex}
                        onChange={(event) =>
                          updateValue(i, { fieldIndex: Number(event.target.value) })
                        }
                        style={{ flex: 1, minWidth: 0 }}
                      >
                        {fields.map((field, fi) => (
                          <option key={field.colIndex} value={fi}>
                            {field.label}
                          </option>
                        ))}
                      </select>
                      <select
                        value={spec.agg}
                        onChange={(event) =>
                          updateValue(i, { agg: event.target.value as PivotValueSpec['agg'] })
                        }
                      >
                        {(Object.entries(AGG_LABELS) as [PivotValueSpec['agg'], StringKey][]).map(
                          ([agg, labelKey]) => (
                            <option key={agg} value={agg}>
                              {t(labelKey)}
                            </option>
                          ),
                        )}
                      </select>
                    </>
                  )}
                  <select
                    data-tip={t('dlgPivotShowAs')}
                    value={spec.showDataAs ?? ''}
                    onChange={(event) => {
                      const v = event.target.value as PivotShowDataAsOption | ''
                      // '' = normal: remove showDataAs from the spec.
                      updateValue(i, { showDataAs: v === '' ? undefined : v })
                    }}
                  >
                    {SHOW_DATA_AS_LABELS.map(({ value, labelKey }) => (
                      <option key={value} value={value}>
                        {t(labelKey)}
                      </option>
                    ))}
                  </select>
                  <select
                    data-tip={t('dlgPivotValueFilter')}
                    value={valueFilters[i]?.op ?? ''}
                    onChange={(event) => {
                      const op = event.target.value as PivotValueFilterOption['op'] | ''
                      setValueFilters((prev) => {
                        const next = { ...prev }
                        if (op === '') delete next[i]
                        else if (op === 'top') next[i] = { op, count: prev[i]?.count ?? 10 }
                        else if (op === 'greaterThan') next[i] = { op, from: prev[i]?.from ?? 0 }
                        else next[i] = { op, from: prev[i]?.from ?? 0, to: prev[i]?.to ?? 100 }
                        return next
                      })
                    }}
                  >
                    <option value="">{t('dlgPivotFilterNone')}</option>
                    <option value="top">{t('dlgPivotFilterTopN')}</option>
                    <option value="greaterThan">{t('dlgPivotFilterGreaterThan')}</option>
                    <option value="between">{t('dlgPivotFilterBetween')}</option>
                  </select>
                  {valueFilters[i]?.op === 'top' && (
                    <input
                      type="number"
                      data-tip={t('dlgPivotTopNCount')}
                      className="cell-input"
                      style={{ width: 52 }}
                      min={1}
                      value={valueFilters[i].count ?? 10}
                      onChange={(event) => {
                        const count = Number(event.target.value)
                        setValueFilters((prev) => ({ ...prev, [i]: { op: 'top', count } }))
                      }}
                    />
                  )}
                  {valueFilters[i]?.op === 'greaterThan' && (
                    <input
                      type="number"
                      data-tip={t('dlgPivotGreaterThanValue')}
                      className="cell-input"
                      style={{ width: 64 }}
                      value={valueFilters[i].from ?? 0}
                      onChange={(event) => {
                        const from = Number(event.target.value)
                        setValueFilters((prev) => ({ ...prev, [i]: { op: 'greaterThan', from } }))
                      }}
                    />
                  )}
                  {valueFilters[i]?.op === 'between' && (
                    <>
                      <input
                        type="number"
                        data-tip={t('dlgPivotBetweenFrom')}
                        className="cell-input"
                        style={{ width: 56 }}
                        value={valueFilters[i].from ?? 0}
                        onChange={(event) => {
                          const from = Number(event.target.value)
                          setValueFilters((prev) => ({
                            ...prev,
                            [i]: { op: 'between', from, to: prev[i]?.to ?? 100 },
                          }))
                        }}
                      />
                      <input
                        type="number"
                        data-tip={t('dlgPivotBetweenTo')}
                        className="cell-input"
                        style={{ width: 56 }}
                        value={valueFilters[i].to ?? 100}
                        onChange={(event) => {
                          const to = Number(event.target.value)
                          setValueFilters((prev) => ({
                            ...prev,
                            [i]: { op: 'between', from: prev[i]?.from ?? 0, to },
                          }))
                        }}
                      />
                    </>
                  )}
                  {values.length > 1 && (
                    <button
                      type="button"
                      className="secondary"
                      style={{ padding: '1px 6px' }}
                      onClick={() => removeValue(i)}
                    >
                      ✕
                    </button>
                  )}
                </div>
              ))}
              {values.length < 8 && (
                <span style={{ display: 'inline-flex', gap: 8 }}>
                  <button
                    type="button"
                    className="secondary"
                    style={{ fontSize: 12, marginTop: 2 }}
                    onClick={addValue}
                  >
                    {t('dlgPivotAddValue')}
                  </button>
                  <button
                    type="button"
                    className="secondary"
                    style={{ fontSize: 12, marginTop: 2 }}
                    onClick={addCalculatedValue}
                  >
                    {t('dlgPivotAddCalc')}
                  </button>
                </span>
              )}
            </fieldset>
            <label>
              {t('dlgPivotTargetCell')}
              <input
                type="text"
                className="cell-input"
                placeholder={t('dlgPivotTargetPlaceholder')}
                value={targetCell}
                disabled={mode === 'edit'}
                data-tip={mode === 'edit' ? t('dlgPivotTargetLocked') : undefined}
                onChange={(event) => setTargetCell(event.target.value.toUpperCase())}
                style={{ width: 80 }}
              />
            </label>
            <p className="dialog-note dialog-span">{t('dlgPivotNote')}</p>
          </div>
        )}
        {error && (
          <p className="dialog-note" role="alert" style={{ color: '#c00' }}>
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          {!noFields && (
            <button
              className="primary-action"
              onClick={() => {
                // Row/column dimensions disallow duplicate fields (repeated
                // levels are meaningless and OOXML forbids them), and one field
                // cannot be both a row and a column dimension.
                if (new Set(rowFieldIndices).size !== rowFieldIndices.length) {
                  setError(t('dlgPivotErrDupRowField'))
                  return
                }
                if (new Set(colFieldIndices).size !== colFieldIndices.length) {
                  setError(t('dlgPivotErrDupColField'))
                  return
                }
                if (colFieldIndices.some((index) => rowFieldIndices.includes(index))) {
                  setError(t('dlgPivotErrRowColOverlap'))
                  return
                }
                if (colFieldIndices.length > 0 && values.length !== 1) {
                  setError(t('dlgPivotErrColNeedsSingleValue'))
                  return
                }
                if (
                  values.some(
                    (spec) =>
                      spec.formula !== undefined &&
                      ((spec.calcName ?? '').trim() === '' || spec.formula.trim() === ''),
                  )
                ) {
                  setError(t('dlgPivotErrCalcFieldIncomplete'))
                  return
                }
                // Grouping/filters only carry fields still on the row/column
                // dimensions (leftovers from field swaps are ignored).
                const axisFieldSet = new Set([...rowFieldIndices, ...colFieldIndices])
                const groupings = Object.entries(fieldGroupings)
                  .map(([fieldIdx, rule]) => ({ fieldIndex: Number(fieldIdx), rule }))
                  .filter(({ fieldIndex }) => axisFieldSet.has(fieldIndex))
                if (
                  groupings.some(
                    ({ rule }) =>
                      rule.kind === 'range' &&
                      (!Number.isFinite(rule.rangeStep) || rule.rangeStep <= 0),
                  )
                ) {
                  setError(t('dlgPivotErrRangeStep'))
                  return
                }
                const labelFilterEntries = Object.entries(labelFilters)
                  .map(([fieldIdx, rule]) => ({ fieldIndex: Number(fieldIdx), rule }))
                  .filter(({ fieldIndex }) => axisFieldSet.has(fieldIndex))
                if (labelFilterEntries.some(({ rule }) => rule.value.trim() === '')) {
                  setError(t('dlgPivotErrLabelFilterText'))
                  return
                }
                const valueFilterEntries = Object.entries(valueFilters)
                  .map(([valueIdx, rule]) => ({ valueIndex: Number(valueIdx), rule }))
                  .filter(({ valueIndex }) => valueIndex < values.length)
                // The value filter applies to the level-1 row field, is mutually
                // exclusive with its label filter, and is at most one.
                if (valueFilterEntries.length > 1) {
                  setError(t('dlgPivotErrOneValueFilter'))
                  return
                }
                if (
                  valueFilterEntries.length === 1 &&
                  labelFilterEntries.some(({ fieldIndex }) => fieldIndex === rowFieldIndices[0])
                ) {
                  setError(t('dlgPivotErrFilterConflict'))
                  return
                }
                const failure = onCreate({
                  sourceRange,
                  rowFieldIndices,
                  colFieldIndices,
                  groupings,
                  labelFilters: labelFilterEntries,
                  valueFilters: valueFilterEntries,
                  values,
                  targetCell: targetCell.trim() || 'A1',
                })
                if (failure !== null) setError(failure)
                else onClose()
              }}
            >
              {mode === 'edit' ? t('dlgApply') : t('dlgCreate')}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
