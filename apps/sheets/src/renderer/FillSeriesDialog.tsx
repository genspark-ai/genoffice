import { useState } from 'react'

import {
  type FillSeriesDateUnit,
  type FillSeriesDirection,
  type FillSeriesOptions,
  type FillSeriesType,
  parseSeriesNumber,
} from './fill-series'
import type { FillSeriesContext } from './fill-series-apply'
import { type StringKey, useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

const DIRECTIONS: readonly [FillSeriesDirection, string][] = [
  ['rows', 'dlgFsRows'],
  ['columns', 'dlgFsColumns'],
]
const TYPES: readonly [FillSeriesType, string][] = [
  ['linear', 'dlgFsLinear'],
  ['growth', 'dlgFsGrowth'],
  ['date', 'dlgFsDate'],
  ['autofill', 'dlgFsAutoFill'],
]
const DATE_UNITS: readonly [FillSeriesDateUnit, string][] = [
  ['day', 'dlgFsDay'],
  ['weekday', 'dlgFsWeekday'],
  ['month', 'dlgFsMonth'],
  ['year', 'dlgFsYear'],
]

/// Excel's Home > Fill > Series… dialog.
export function FillSeriesDialog({
  context,
  defaultType,
  onApply,
  onClose,
}: {
  readonly context: FillSeriesContext
  readonly defaultType: FillSeriesType
  /// Returns a user-facing error, or null when the fill was applied.
  readonly onApply: (options: FillSeriesOptions) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [direction, setDirection] = useState<FillSeriesDirection>(context.direction)
  const [type, setType] = useState<FillSeriesType>(defaultType)
  const [dateUnit, setDateUnit] = useState<FillSeriesDateUnit>('day')
  const [trend, setTrend] = useState(false)
  const [step, setStep] = useState('1')
  const [stop, setStop] = useState('')
  const [error, setError] = useState<string | null>(null)

  const stepValue = parseSeriesNumber(step, false)
  const stopValue = parseSeriesNumber(stop, type === 'date', context.date1904)
  const trendable = type === 'linear' || type === 'growth'
  const valid =
    (trend && trendable) || (stepValue !== null && (stop.trim() === '' || stopValue !== null))

  const apply = (): void => {
    const failure = onApply({
      direction,
      type,
      dateUnit,
      trend: trend && trendable,
      step: stepValue ?? 1,
      stop: trend && trendable ? null : stopValue,
      date1904: context.date1904,
    })
    if (failure) setError(failure)
    else onClose()
  }

  const modal = useModalDialog(onClose)
  const radioGroup = <T extends string>(
    name: string,
    label: string,
    choices: readonly [T, string][],
    value: T,
    onChange: (next: T) => void,
    disabled = false,
  ): React.JSX.Element => (
    <fieldset className="dialog-fieldset" disabled={disabled}>
      <legend>{label}</legend>
      {choices.map(([choice, key]) => (
        <label key={choice} className="dialog-check">
          <input
            type="radio"
            name={name}
            checked={value === choice}
            onChange={() => onChange(choice)}
          />
          {t(key as StringKey)}
        </label>
      ))}
    </fieldset>
  )

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog fill-series-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgFillSeriesTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgFillSeriesTitle')}</header>
        <section className="dialog-body">
          <div className="fill-series-groups">
            {radioGroup('fs-direction', t('dlgFsSeriesIn'), DIRECTIONS, direction, setDirection)}
            {radioGroup('fs-type', t('dlgFsType'), TYPES, type, setType)}
            {radioGroup(
              'fs-unit',
              t('dlgFsDateUnit'),
              DATE_UNITS,
              dateUnit,
              setDateUnit,
              type !== 'date',
            )}
          </div>
          <div className="dialog-grid">
            <label className="dialog-check dialog-span">
              <input
                type="checkbox"
                checked={trend && trendable}
                disabled={!trendable}
                onChange={(event) => setTrend(event.target.checked)}
              />
              {t('dlgFsTrend')}
            </label>
            <label>
              {t('dlgFsStepValue')}
              <input
                autoFocus
                value={step}
                disabled={trend && trendable}
                onChange={(event) => setStep(event.target.value)}
              />
            </label>
            <label>
              {t('dlgFsStopValue')}
              <input
                value={stop}
                disabled={trend && trendable}
                onChange={(event) => setStop(event.target.value)}
              />
            </label>
          </div>
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" disabled={!valid} onClick={apply}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
