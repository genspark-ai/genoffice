import { useState } from 'react'

import { Dropdown } from '@genoffice/ui'
import { BooleanNumber } from '@univerjs/core'
import type { IFilterColumn } from '@univerjs/preset-sheets-filter'

import { t, useI18n, type StringKey } from './i18n/locale'

/// Excel's custom auto-filter ("Advanced" in Data → Sort & Filter), minimal:
/// one column, up to two conditions joined by AND/OR. OK hands the criteria
/// to App, which lands them through the same path as the AI op
/// set_filter_criteria.

/// Mirrors both the gateway's custom-filter operators and Univer's
/// CustomFilterOperator string values (whose .d.ts declaration is missing,
/// so the enum itself would type as `any`).
export type AdvancedFilterOperator =
  'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual'

export interface AdvancedFilterCondition {
  readonly operator: AdvancedFilterOperator
  readonly val: string
}

/// A column choice: colId is the 0-based offset inside the filter range
/// (the same colId the gateway serializes).
export interface AdvancedFilterColumn {
  readonly colId: number
  readonly label: string
}

export interface AdvancedFilterCriteria {
  readonly colId: number
  readonly and: boolean
  readonly filters: readonly AdvancedFilterCondition[]
}

const OPERATOR_OPTIONS: readonly {
  readonly value: AdvancedFilterOperator
  readonly labelKey: StringKey
}[] = [
  { value: 'equal', labelKey: 'dlgAdvFilterOpEqual' },
  { value: 'notEqual', labelKey: 'dlgAdvFilterOpNotEqual' },
  { value: 'greaterThan', labelKey: 'dlgAdvFilterOpGreaterThan' },
  { value: 'greaterThanOrEqual', labelKey: 'dlgAdvFilterOpGreaterThanOrEqual' },
  { value: 'lessThan', labelKey: 'dlgAdvFilterOpLessThan' },
  { value: 'lessThanOrEqual', labelKey: 'dlgAdvFilterOpLessThanOrEqual' },
]

/// Builds the Univer `customFilters` payload from the dialog's conditions.
/// Univer evaluates an explicit `equal` operator numerically only, so it is
/// dropped: a missing operator is the OOXML default (equality, with text
/// matching) — exactly how the gateway serializes `equal` too.
export function buildCustomFilters(
  and: boolean,
  conditions: readonly AdvancedFilterCondition[],
): NonNullable<IFilterColumn['customFilters']> {
  const [first, second] = conditions
  if (first === undefined) throw new Error(t('dlgAdvFilterNeedCondition'))
  if (conditions.length > 2) throw new Error(t('dlgAdvFilterMaxTwo'))
  return {
    // `and` only matters (and only serializes) with two conditions.
    ...(and && second !== undefined ? { and: BooleanNumber.TRUE } : {}),
    customFilters:
      second === undefined
        ? [toUniverCustomFilter(first)]
        : [toUniverCustomFilter(first), toUniverCustomFilter(second)],
  }
}

function toUniverCustomFilter(condition: AdvancedFilterCondition): {
  val: string
  operator?: AdvancedFilterOperator
} {
  return condition.operator === 'equal'
    ? { val: condition.val }
    : { val: condition.val, operator: condition.operator }
}

export function AdvancedFilterDialog({
  columns,
  onApply,
  onClose,
}: {
  readonly columns: readonly AdvancedFilterColumn[]
  /// Returns an error message, or null on success.
  readonly onApply: (criteria: AdvancedFilterCriteria) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [colId, setColId] = useState(columns[0]?.colId ?? 0)
  const [operator1, setOperator1] = useState<AdvancedFilterOperator>('equal')
  const [value1, setValue1] = useState('')
  const [operator2, setOperator2] = useState<AdvancedFilterOperator>('equal')
  const [value2, setValue2] = useState('')
  const [and, setAnd] = useState(true)
  const [error, setError] = useState<string | null>(null)

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgAdvFilterTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgAdvFilterTitle')}</header>
        {columns.length === 0 ? (
          <p className="dialog-note">{t('dlgAdvFilterNoColumns')}</p>
        ) : (
          <div className="dialog-grid">
            <label className="dialog-span">
              {t('dlgAdvFilterColumn')}
              <Dropdown
                ariaLabel={t('dlgAdvFilterColumn')}
                value={String(colId)}
                options={columns.map((column) => ({
                  value: String(column.colId),
                  label: column.label,
                }))}
                onPick={(v) => setColId(Number(v))}
              />
            </label>
            <label>
              {t('dlgAdvFilterCondition1')}
              <Dropdown
                ariaLabel={t('dlgAdvFilterCondition1')}
                value={operator1}
                options={OPERATOR_OPTIONS.map((option) => ({
                  value: option.value,
                  label: t(option.labelKey),
                }))}
                onPick={setOperator1}
              />
            </label>
            <label>
              {t('dlgAdvFilterValue')}
              <input value={value1} onChange={(event) => setValue1(event.target.value)} />
            </label>
            <label className="dialog-check">
              <input
                type="radio"
                name="advanced-filter-join"
                checked={and}
                onChange={() => setAnd(true)}
              />
              {t('dlgAdvFilterAnd')}
            </label>
            <label className="dialog-check">
              <input
                type="radio"
                name="advanced-filter-join"
                checked={!and}
                onChange={() => setAnd(false)}
              />
              {t('dlgAdvFilterOr')}
            </label>
            <label>
              {t('dlgAdvFilterCondition2')}
              <Dropdown
                ariaLabel={t('dlgAdvFilterCondition2')}
                value={operator2}
                options={OPERATOR_OPTIONS.map((option) => ({
                  value: option.value,
                  label: t(option.labelKey),
                }))}
                onPick={setOperator2}
              />
            </label>
            <label>
              {t('dlgAdvFilterValue')}
              <input value={value2} onChange={(event) => setValue2(event.target.value)} />
            </label>
            <p className="dialog-note dialog-span">{t('dlgAdvFilterNote')}</p>
          </div>
        )}
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          {columns.length > 0 && (
            <button
              className="primary-action"
              onClick={() => {
                if (value1 === '') {
                  setError(t('dlgAdvFilterNeedValue1'))
                  return
                }
                const failure = onApply({
                  colId,
                  and,
                  filters: [
                    { operator: operator1, val: value1 },
                    ...(value2 === '' ? [] : [{ operator: operator2, val: value2 }]),
                  ],
                })
                setError(failure)
                if (failure === null) onClose()
              }}
            >
              {t('dlgOk')}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
