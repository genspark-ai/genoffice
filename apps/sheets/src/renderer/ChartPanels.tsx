import { useState } from 'react'

import { transposeChartSeries, type ChartVisualState } from '../domain/chart-visual'
import { ColorDropdown } from './ColorDropdown'
import { useI18n, type StringKey, type TFunc } from './i18n/locale'
import type { ChartEditData, ChartElementRef, ChartVectorRead } from './WorkbookVisuals'

const LEGEND_OPTIONS: readonly (readonly [string, StringKey])[] = [
  ['right', 'appLegendRight'],
  ['bottom', 'appLegendBottom'],
  ['top', 'appLegendTop'],
  ['left', 'appLegendLeft'],
  ['none', 'appNone'],
]
const LABEL_OPTIONS: readonly (readonly [string, StringKey])[] = [
  ['none', 'appNone'],
  ['value', 'appLabelValue'],
  ['percent', 'appLabelPercent'],
  ['category-percent', 'appLabelCategoryPercent'],
]

function chartFamily(chart: ChartVisualState): {
  isPie: boolean
  isDoughnut: boolean
  isBar: boolean
  hasAxes: boolean
} {
  const types = chart.chartTypes
  const isDoughnut = types.includes('doughnutChart')
  const isPie = isDoughnut || types.includes('pieChart')
  return {
    isPie,
    isDoughnut,
    isBar: types.includes('barChart'),
    hasAxes: !isPie && types.length > 0,
  }
}

function elementLabel(t: TFunc, element: ChartElementRef | null, chart: ChartVisualState): string {
  if (!element) return t('appChartArea')
  if (element.kind === 'series' || element.kind === 'point') {
    const series = chart.series[element.seriesIndex]
    const name = series?.name || t('appSeriesN', { n: element.seriesIndex + 1 })
    if (element.kind === 'series') return t('appSeriesQuoted', { name })
    const category = series?.categories[element.pointIndex] ?? `#${element.pointIndex + 1}`
    return t('appDataPointQuoted', { category })
  }
  return t(
    (
      {
        chart: 'appChartArea',
        title: 'appChartTitleEl',
        legend: 'appLegendEl',
        'value-axis': 'appValueAxis',
        'category-axis': 'appCategoryAxis',
      } as const
    )[element.kind],
  )
}

/// Number field that commits on blur/Enter; empty commits null (= auto).
function BoundInput({
  label,
  value,
  onCommit,
}: {
  readonly label: string
  readonly value: number | undefined
  readonly onCommit: (next: number | null) => void
}): React.JSX.Element {
  const { t } = useI18n()
  return (
    <label>
      {label}
      <input
        type="number"
        key={value ?? 'auto'}
        defaultValue={value ?? ''}
        placeholder={t('appAuto')}
        onBlur={(event) => {
          const raw = event.target.value.trim()
          const parsed = Number(raw)
          if (raw === '') {
            if (value !== undefined) onCommit(null)
          } else if (Number.isFinite(parsed) && parsed !== value) {
            onCommit(parsed)
          }
        }}
        onKeyDown={(event) => {
          if (event.key === 'Enter') (event.target as HTMLInputElement).blur()
        }}
      />
    </label>
  )
}

/// Range that commits on release (pointer up / blur), not per tick — a drag
/// must land as one undo step, not dozens.
function BoundRange({
  label,
  min,
  max,
  step,
  value,
  onCommit,
}: {
  readonly label: string
  readonly min: number
  readonly max: number
  readonly step: number
  readonly value: number
  readonly onCommit: (next: number) => void
}): React.JSX.Element {
  const commit = (raw: string): void => {
    const next = Number(raw)
    if (Number.isFinite(next) && next !== value) onCommit(next)
  }
  return (
    <label>
      {label}
      <input
        type="range"
        min={min}
        max={max}
        step={step}
        key={value}
        defaultValue={value}
        onPointerUp={(event) => commit(event.currentTarget.value)}
        onBlur={(event) => commit(event.currentTarget.value)}
      />
    </label>
  )
}

/// Chart format task pane, scoped by the selected chart element.
export function ChartFormatPane({
  chart,
  element,
  onEdit,
  onClose,
}: {
  readonly chart: ChartVisualState
  readonly element: ChartElementRef | null
  readonly onEdit: (edit: ChartEditData) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const { isPie, isDoughnut, isBar, hasAxes } = chartFamily(chart)
  const primary = chart.series[0]
  return (
    <div className="chart-format-pane">
      <header>
        <span>{t('appFormatHeader', { target: elementLabel(t, element, chart) })}</span>
        <button data-tip={t('appClose')} aria-label={t('appClose')} onClick={onClose}>
          ✕
        </button>
      </header>
      {(element?.kind === 'series' || element?.kind === 'point') && (
        <div className="pane-section">
          <strong>{element.kind === 'series' ? t('appSeries') : t('appDataPoint')}</strong>
          <label>
            {t('appFillColor')}
            <ColorDropdown
              label={t('appFillColor')}
              value={
                (element.kind === 'point'
                  ? chart.series[element.seriesIndex]?.pointColors?.find(
                      (entry) => entry.index === element.pointIndex,
                    )?.color
                  : undefined) ??
                chart.series[element.seriesIndex]?.color ??
                '#4472c4'
              }
              onPick={(color) => {
                if (!color) return
                if (element.kind === 'series') {
                  onEdit({ seriesColors: { [String(element.seriesIndex)]: color } })
                } else {
                  onEdit({
                    pointColors: {
                      [String(element.seriesIndex)]: { [String(element.pointIndex)]: color },
                    },
                  })
                }
              }}
            />
          </label>
          {isPie && element.kind === 'point' && element.seriesIndex === 0 && (
            <BoundRange
              label={t('appPointExplosion')}
              min={0}
              max={100}
              step={5}
              value={
                primary?.pointExplosions?.find((entry) => entry.index === element.pointIndex)
                  ?.pct ??
                primary?.explosionPct ??
                0
              }
              onCommit={(next) =>
                onEdit({
                  pointExplosions: { [String(element.pointIndex)]: next },
                })
              }
            />
          )}
        </div>
      )}
      <div className="pane-section">
        <strong>{t('appTitle')}</strong>
        <input
          type="text"
          key={chart.title}
          defaultValue={chart.title}
          maxLength={255}
          onBlur={(event) => {
            const next = event.target.value
            if (next !== chart.title) onEdit({ title: next })
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') (event.target as HTMLInputElement).blur()
          }}
        />
      </div>
      <div className="pane-section">
        <strong>{t('appLegendAndLabels')}</strong>
        <label>
          {t('appLegendPosition')}
          <select
            value={chart.legend ?? 'right'}
            onChange={(event) =>
              onEdit({ legend: event.target.value as NonNullable<ChartEditData['legend']> })
            }
          >
            {LEGEND_OPTIONS.map(([value, key]) => (
              <option key={value} value={value}>
                {t(key)}
              </option>
            ))}
          </select>
        </label>
        <label>
          {t('appDataLabels')}
          <select
            value={chart.dataLabels ?? 'none'}
            onChange={(event) =>
              onEdit({ dataLabels: event.target.value as NonNullable<ChartEditData['dataLabels']> })
            }
          >
            {LABEL_OPTIONS.filter(
              ([value]) => isPie || (value !== 'percent' && value !== 'category-percent'),
            ).map(([value, key]) => (
              <option key={value} value={value}>
                {t(key)}
              </option>
            ))}
          </select>
        </label>
        {(isBar || isPie) && chart.dataLabels !== undefined && chart.dataLabels !== 'none' && (
          <label>
            {t('appLabelPosition')}
            <select
              value={chart.dataLabelPosition ?? ''}
              onChange={(event) => {
                if (event.target.value) {
                  onEdit({
                    dataLabelPosition: event.target.value as NonNullable<
                      ChartEditData['dataLabelPosition']
                    >,
                  })
                }
              }}
            >
              <option value="" disabled hidden>
                {t('appAuto')}
              </option>
              <option value="outside-end">{t('appOutsideEnd')}</option>
              <option value="inside-end">{t('appInsideEnd')}</option>
              <option value="center">{t('appCenter')}</option>
            </select>
          </label>
        )}
        {chart.dataLabels !== undefined && chart.dataLabels !== 'none' && (
          <label>
            {t('appLabelNumberFormat')}
            <input
              type="text"
              key={chart.dataLabelFormat ?? ''}
              defaultValue={chart.dataLabelFormat ?? ''}
              placeholder={t('appLabelFormatPlaceholder')}
              maxLength={64}
              onBlur={(event) => {
                const next = event.target.value.trim()
                if (next && next !== chart.dataLabelFormat) onEdit({ dataLabelFormat: next })
              }}
              onKeyDown={(event) => {
                if (event.key === 'Enter') (event.target as HTMLInputElement).blur()
              }}
            />
          </label>
        )}
      </div>
      {hasAxes && (
        <div className="pane-section">
          <strong>{t('appValueAxis')}</strong>
          <label>
            {t('appGridlines')}
            <input
              type="checkbox"
              checked={chart.gridlines !== false}
              onChange={(event) => onEdit({ gridlines: event.target.checked })}
            />
          </label>
          <BoundInput
            label={t('appAxisMin')}
            value={chart.valueAxis?.min}
            onCommit={(min) => onEdit({ valueAxis: { min } })}
          />
          <BoundInput
            label={t('appAxisMax')}
            value={chart.valueAxis?.max}
            onCommit={(max) => onEdit({ valueAxis: { max } })}
          />
        </div>
      )}
      {isBar && (
        <div className="pane-section">
          <strong>{t('appBarSection')}</strong>
          <BoundRange
            label={t('appGapWidth')}
            min={0}
            max={500}
            step={10}
            value={chart.gapWidthPct ?? 150}
            onCommit={(next) => onEdit({ gapWidthPct: next })}
          />
        </div>
      )}
      {isPie && (
        <div className="pane-section">
          <strong>{isDoughnut ? t('appDoughnutSection') : t('appChartPie')}</strong>
          <BoundRange
            label={t('appPieExplosion')}
            min={0}
            max={100}
            step={5}
            value={primary?.explosionPct ?? 0}
            onCommit={(next) => onEdit({ explosionPct: next })}
          />
          {isDoughnut && (
            <BoundRange
              label={t('appHoleSize')}
              min={10}
              max={90}
              step={5}
              value={chart.holeSizePct ?? 50}
              onCommit={(next) => onEdit({ holeSizePct: next })}
            />
          )}
        </div>
      )}
    </div>
  )
}

interface SeriesRow {
  name: string
  valuesRange: string
  categoriesRange: string
  /// Index into the chart's current series; null for a freshly added row.
  source: number | null
}

/// Excel's Select Data Source: rename / repoint / add / remove / reorder
/// series, plus switch row-column. Applies as one full series replacement.
export function SelectDataDialog({
  chart,
  supported,
  readVector,
  onApply,
  onClose,
}: {
  readonly chart: ChartVisualState
  /// Single-plot chart families only; scatter/combos fail closed at save.
  readonly supported: boolean
  readonly readVector?: ((range: string) => Promise<ChartVectorRead>) | undefined
  readonly onApply: (edit: ChartEditData) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [rows, setRows] = useState<SeriesRow[]>(
    chart.series.map((series, index) => ({
      name: series.name,
      valuesRange: '',
      categoriesRange: '',
      source: index,
    })),
  )
  const [error, setError] = useState<string | null>(null)
  const [isBusy, setIsBusy] = useState(false)

  const patchRow = (index: number, patch: Partial<SeriesRow>): void =>
    setRows((previous) => previous.map((row, at) => (at === index ? { ...row, ...patch } : row)))
  const moveRow = (index: number, delta: -1 | 1): void =>
    setRows((previous) => {
      const next = [...previous]
      const target = index + delta
      const row = next[index]
      const other = next[target]
      if (!row || !other) return previous
      next[index] = other
      next[target] = row
      return next
    })

  async function buildSeriesSet(): Promise<NonNullable<ChartEditData['seriesSet']>> {
    const entries: NonNullable<ChartEditData['seriesSet']> = []
    for (const [at, row] of rows.entries()) {
      const original = row.source === null ? undefined : chart.series[row.source]
      const name = row.name.trim() || original?.name || t('appSeriesN', { n: at + 1 })
      const valuesRange = row.valuesRange.trim()
      const categoriesRange = row.categoriesRange.trim()
      if (!original && !valuesRange) {
        throw new Error(t('appSeriesNeedsValues', { name }))
      }
      if ((valuesRange || categoriesRange) && !readVector) {
        throw new Error(t('appChartNoRepoint'))
      }
      let values = original?.values ?? []
      let valuesRef = original?.valuesRef
      if (valuesRange && readVector) {
        const read = await readVector(valuesRange)
        values = read.vector.map((value) => {
          const numeric = typeof value === 'number' ? value : Number(String(value ?? '').trim())
          return Number.isFinite(numeric) ? numeric : 0
        })
        valuesRef = read.ref
      }
      let categories = original?.categories
      let categoriesRef = original?.categoriesRef
      if (categoriesRange && readVector) {
        const read = await readVector(categoriesRange)
        categories = read.vector.map((value) => String(value ?? '').slice(0, 255))
        categoriesRef = read.ref
      }
      entries.push({
        name: name.slice(0, 255),
        values,
        ...(categories === undefined ? {} : { categories }),
        ...(valuesRef === undefined ? {} : { valuesRef }),
        ...(categoriesRef === undefined ? {} : { categoriesRef }),
        ...(original?.color === undefined ? {} : { color: original.color }),
      })
    }
    if (entries.length === 0) throw new Error(t('appKeepOneSeries'))
    return entries
  }

  function switchRowColumn(): void {
    const seriesSet = transposeChartSeries(chart.series, (n) => t('appSeriesN', { n }))
    if (!seriesSet) {
      setError(t('appNoCategoriesToSwitch'))
      return
    }
    onApply({ seriesSet })
    onClose()
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog select-data-dialog"
        role="dialog"
        aria-label={t('appSelectDataSource')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appSelectDataSource')}</header>
        <div className="dialog-body">
          {!supported && <p>{t('appComboNoSelectData')}</p>}
          {supported && (
            <>
              <div className="select-data-row select-data-head">
                <span>{t('appSeriesName')}</span>
                <span>{t('appValuesRange')}</span>
                <span>{t('appCategoriesRange')}</span>
                <span />
              </div>
              <div className="select-data-rows">
                {rows.map((row, index) => (
                  <div className="select-data-row" key={index}>
                    <input
                      value={row.name}
                      placeholder={t('appSeriesN', { n: index + 1 })}
                      onChange={(event) => patchRow(index, { name: event.target.value })}
                    />
                    <input
                      value={row.valuesRange}
                      placeholder={
                        row.source !== null
                          ? (chart.series[row.source]?.valuesRef ?? t('appKeepUnchanged'))
                          : t('appValuesRangeExample')
                      }
                      onChange={(event) => patchRow(index, { valuesRange: event.target.value })}
                    />
                    <input
                      value={row.categoriesRange}
                      placeholder={
                        row.source !== null
                          ? (chart.series[row.source]?.categoriesRef ?? t('appKeepUnchanged'))
                          : t('appCategoriesRangeExample')
                      }
                      onChange={(event) => patchRow(index, { categoriesRange: event.target.value })}
                    />
                    <span>
                      <button
                        data-tip={t('appMoveUp')}
                        aria-label={t('appMoveUp')}
                        disabled={index === 0}
                        onClick={() => moveRow(index, -1)}
                      >
                        ↑
                      </button>
                      <button
                        data-tip={t('appMoveDown')}
                        aria-label={t('appMoveDown')}
                        disabled={index === rows.length - 1}
                        onClick={() => moveRow(index, 1)}
                      >
                        ↓
                      </button>
                      <button
                        data-tip={t('appRemoveSeries')}
                        aria-label={t('appRemoveSeries')}
                        onClick={() =>
                          setRows((previous) => previous.filter((_, at) => at !== index))
                        }
                      >
                        ✕
                      </button>
                    </span>
                  </div>
                ))}
              </div>
              <div className="select-data-actions">
                <button
                  disabled={rows.length >= 24}
                  onClick={() =>
                    setRows((previous) => [
                      ...previous,
                      { name: '', valuesRange: '', categoriesRange: '', source: null },
                    ])
                  }
                >
                  {t('appAddSeries')}
                </button>
                <button disabled={chart.series.length === 0} onClick={switchRowColumn}>
                  {t('appSwitchRowColumn')}
                </button>
              </div>
            </>
          )}
          {error && <div className="chart-editor-error">{error}</div>}
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button
            className="primary"
            disabled={!supported || isBusy || rows.length === 0}
            onClick={() => {
              setIsBusy(true)
              setError(null)
              buildSeriesSet()
                .then((seriesSet) => {
                  onApply({ seriesSet })
                  onClose()
                })
                .catch((reason: unknown) => {
                  setError(reason instanceof Error ? reason.message : t('appCannotReadDataRange'))
                })
                .finally(() => setIsBusy(false))
            }}
          >
            {isBusy ? t('appReading') : t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
