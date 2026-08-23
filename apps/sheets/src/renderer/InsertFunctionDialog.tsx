import { useMemo, useState } from 'react'

import { Dropdown } from '@genoffice/ui'

import { useI18n, type StringKey } from './i18n/locale'

/// Excel's Insert Function, minimal: browse/search the catalog, read the
/// syntax, finish the formula in the dialog, apply to the active cell.

interface FunctionSpec {
  readonly name: string
  /// Stable English id; displayed through CATEGORY_LABELS.
  readonly category: string
  readonly syntax: string
  readonly descKey: StringKey
}

const FUNCTION_CATALOG: readonly FunctionSpec[] = [
  { name: 'SUM', category: 'Math', syntax: 'SUM(number1, [number2], …)', descKey: 'dlgFnDescSum' },
  {
    name: 'SUMIF',
    category: 'Math',
    syntax: 'SUMIF(range, criteria, [sum_range])',
    descKey: 'dlgFnDescSumif',
  },
  {
    name: 'SUMIFS',
    category: 'Math',
    syntax: 'SUMIFS(sum_range, criteria_range1, criteria1, …)',
    descKey: 'dlgFnDescSumifs',
  },
  {
    name: 'SUMPRODUCT',
    category: 'Math',
    syntax: 'SUMPRODUCT(array1, [array2], …)',
    descKey: 'dlgFnDescSumproduct',
  },
  {
    name: 'SUBTOTAL',
    category: 'Math',
    syntax: 'SUBTOTAL(function_num, ref1, …)',
    descKey: 'dlgFnDescSubtotal',
  },
  {
    name: 'ROUND',
    category: 'Math',
    syntax: 'ROUND(number, num_digits)',
    descKey: 'dlgFnDescRound',
  },
  {
    name: 'ROUNDUP',
    category: 'Math',
    syntax: 'ROUNDUP(number, num_digits)',
    descKey: 'dlgFnDescRoundup',
  },
  {
    name: 'ROUNDDOWN',
    category: 'Math',
    syntax: 'ROUNDDOWN(number, num_digits)',
    descKey: 'dlgFnDescRounddown',
  },
  { name: 'ABS', category: 'Math', syntax: 'ABS(number)', descKey: 'dlgFnDescAbs' },
  { name: 'INT', category: 'Math', syntax: 'INT(number)', descKey: 'dlgFnDescInt' },
  { name: 'MOD', category: 'Math', syntax: 'MOD(number, divisor)', descKey: 'dlgFnDescMod' },
  { name: 'POWER', category: 'Math', syntax: 'POWER(number, power)', descKey: 'dlgFnDescPower' },
  { name: 'SQRT', category: 'Math', syntax: 'SQRT(number)', descKey: 'dlgFnDescSqrt' },
  { name: 'RAND', category: 'Math', syntax: 'RAND()', descKey: 'dlgFnDescRand' },
  {
    name: 'RANDBETWEEN',
    category: 'Math',
    syntax: 'RANDBETWEEN(bottom, top)',
    descKey: 'dlgFnDescRandbetween',
  },
  {
    name: 'AVERAGE',
    category: 'Statistical',
    syntax: 'AVERAGE(number1, [number2], …)',
    descKey: 'dlgFnDescAverage',
  },
  {
    name: 'AVERAGEIF',
    category: 'Statistical',
    syntax: 'AVERAGEIF(range, criteria, [average_range])',
    descKey: 'dlgFnDescAverageif',
  },
  {
    name: 'COUNT',
    category: 'Statistical',
    syntax: 'COUNT(value1, [value2], …)',
    descKey: 'dlgFnDescCount',
  },
  {
    name: 'COUNTA',
    category: 'Statistical',
    syntax: 'COUNTA(value1, [value2], …)',
    descKey: 'dlgFnDescCounta',
  },
  {
    name: 'COUNTIF',
    category: 'Statistical',
    syntax: 'COUNTIF(range, criteria)',
    descKey: 'dlgFnDescCountif',
  },
  {
    name: 'COUNTIFS',
    category: 'Statistical',
    syntax: 'COUNTIFS(criteria_range1, criteria1, …)',
    descKey: 'dlgFnDescCountifs',
  },
  {
    name: 'MIN',
    category: 'Statistical',
    syntax: 'MIN(number1, [number2], …)',
    descKey: 'dlgFnDescMin',
  },
  {
    name: 'MAX',
    category: 'Statistical',
    syntax: 'MAX(number1, [number2], …)',
    descKey: 'dlgFnDescMax',
  },
  {
    name: 'MEDIAN',
    category: 'Statistical',
    syntax: 'MEDIAN(number1, [number2], …)',
    descKey: 'dlgFnDescMedian',
  },
  {
    name: 'RANK',
    category: 'Statistical',
    syntax: 'RANK(number, ref, [order])',
    descKey: 'dlgFnDescRank',
  },
  { name: 'LARGE', category: 'Statistical', syntax: 'LARGE(array, k)', descKey: 'dlgFnDescLarge' },
  { name: 'SMALL', category: 'Statistical', syntax: 'SMALL(array, k)', descKey: 'dlgFnDescSmall' },
  {
    name: 'IF',
    category: 'Logical',
    syntax: 'IF(logical_test, value_if_true, [value_if_false])',
    descKey: 'dlgFnDescIf',
  },
  {
    name: 'IFERROR',
    category: 'Logical',
    syntax: 'IFERROR(value, value_if_error)',
    descKey: 'dlgFnDescIferror',
  },
  {
    name: 'AND',
    category: 'Logical',
    syntax: 'AND(logical1, [logical2], …)',
    descKey: 'dlgFnDescAnd',
  },
  {
    name: 'OR',
    category: 'Logical',
    syntax: 'OR(logical1, [logical2], …)',
    descKey: 'dlgFnDescOr',
  },
  { name: 'NOT', category: 'Logical', syntax: 'NOT(logical)', descKey: 'dlgFnDescNot' },
  {
    name: 'VLOOKUP',
    category: 'Lookup',
    syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    descKey: 'dlgFnDescVlookup',
  },
  {
    name: 'HLOOKUP',
    category: 'Lookup',
    syntax: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
    descKey: 'dlgFnDescHlookup',
  },
  {
    name: 'INDEX',
    category: 'Lookup',
    syntax: 'INDEX(array, row_num, [column_num])',
    descKey: 'dlgFnDescIndex',
  },
  {
    name: 'MATCH',
    category: 'Lookup',
    syntax: 'MATCH(lookup_value, lookup_array, [match_type])',
    descKey: 'dlgFnDescMatch',
  },
  {
    name: 'CHOOSE',
    category: 'Lookup',
    syntax: 'CHOOSE(index_num, value1, [value2], …)',
    descKey: 'dlgFnDescChoose',
  },
  {
    name: 'CONCATENATE',
    category: 'Text',
    syntax: 'CONCATENATE(text1, [text2], …)',
    descKey: 'dlgFnDescConcatenate',
  },
  { name: 'TEXT', category: 'Text', syntax: 'TEXT(value, format_text)', descKey: 'dlgFnDescText' },
  { name: 'LEFT', category: 'Text', syntax: 'LEFT(text, [num_chars])', descKey: 'dlgFnDescLeft' },
  {
    name: 'RIGHT',
    category: 'Text',
    syntax: 'RIGHT(text, [num_chars])',
    descKey: 'dlgFnDescRight',
  },
  {
    name: 'MID',
    category: 'Text',
    syntax: 'MID(text, start_num, num_chars)',
    descKey: 'dlgFnDescMid',
  },
  { name: 'LEN', category: 'Text', syntax: 'LEN(text)', descKey: 'dlgFnDescLen' },
  { name: 'TRIM', category: 'Text', syntax: 'TRIM(text)', descKey: 'dlgFnDescTrim' },
  { name: 'UPPER', category: 'Text', syntax: 'UPPER(text)', descKey: 'dlgFnDescUpper' },
  { name: 'LOWER', category: 'Text', syntax: 'LOWER(text)', descKey: 'dlgFnDescLower' },
  {
    name: 'SUBSTITUTE',
    category: 'Text',
    syntax: 'SUBSTITUTE(text, old_text, new_text, [instance_num])',
    descKey: 'dlgFnDescSubstitute',
  },
  { name: 'TODAY', category: 'Date & Time', syntax: 'TODAY()', descKey: 'dlgFnDescToday' },
  { name: 'NOW', category: 'Date & Time', syntax: 'NOW()', descKey: 'dlgFnDescNow' },
  {
    name: 'DATE',
    category: 'Date & Time',
    syntax: 'DATE(year, month, day)',
    descKey: 'dlgFnDescDate',
  },
  {
    name: 'YEAR',
    category: 'Date & Time',
    syntax: 'YEAR(serial_number)',
    descKey: 'dlgFnDescYear',
  },
  {
    name: 'MONTH',
    category: 'Date & Time',
    syntax: 'MONTH(serial_number)',
    descKey: 'dlgFnDescMonth',
  },
  { name: 'DAY', category: 'Date & Time', syntax: 'DAY(serial_number)', descKey: 'dlgFnDescDay' },
  {
    name: 'EDATE',
    category: 'Date & Time',
    syntax: 'EDATE(start_date, months)',
    descKey: 'dlgFnDescEdate',
  },
  {
    name: 'PMT',
    category: 'Financial',
    syntax: 'PMT(rate, nper, pv, [fv], [type])',
    descKey: 'dlgFnDescPmt',
  },
  {
    name: 'FV',
    category: 'Financial',
    syntax: 'FV(rate, nper, pmt, [pv], [type])',
    descKey: 'dlgFnDescFv',
  },
  {
    name: 'PV',
    category: 'Financial',
    syntax: 'PV(rate, nper, pmt, [fv], [type])',
    descKey: 'dlgFnDescPv',
  },
  {
    name: 'RATE',
    category: 'Financial',
    syntax: 'RATE(nper, pmt, pv, [fv], [type], [guess])',
    descKey: 'dlgFnDescRate',
  },
  {
    name: 'NPER',
    category: 'Financial',
    syntax: 'NPER(rate, pmt, pv, [fv], [type])',
    descKey: 'dlgFnDescNper',
  },
  {
    name: 'NPV',
    category: 'Financial',
    syntax: 'NPV(rate, value1, [value2], …)',
    descKey: 'dlgFnDescNpv',
  },
  { name: 'IRR', category: 'Financial', syntax: 'IRR(values, [guess])', descKey: 'dlgFnDescIrr' },
]

const CATEGORIES = ['All', ...new Set(FUNCTION_CATALOG.map((spec) => spec.category))]

const CATEGORY_LABELS: Record<string, StringKey> = {
  All: 'dlgFnCatAll',
  Math: 'dlgFnCatMath',
  Statistical: 'dlgFnCatStatistical',
  Logical: 'dlgFnCatLogical',
  Lookup: 'dlgFnCatLookup',
  Text: 'dlgFnCatText',
  'Date & Time': 'dlgFnCatDateTime',
  Financial: 'dlgFnCatFinancial',
}

export function InsertFunctionDialog({
  targetLabel,
  onApply,
  onClose,
  initialCategory,
}: {
  /// A1 label of the destination cell, for the dialog header.
  readonly targetLabel: string
  /// Returns an error message, or null on success.
  readonly onApply: (formula: string) => string | null
  readonly onClose: () => void
  /// Category to open on (the Formulas tab's category buttons pass their own).
  readonly initialCategory?: string
}): React.JSX.Element {
  const { t, lang } = useI18n()
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState(
    initialCategory && CATEGORIES.includes(initialCategory) ? initialCategory : 'All',
  )
  const [picked, setPicked] = useState<FunctionSpec | null>(null)
  const [formula, setFormula] = useState('')
  const [error, setError] = useState<string | null>(null)

  const matches = useMemo(() => {
    const needle = query.trim().toUpperCase()
    return FUNCTION_CATALOG.filter(
      (spec) =>
        (category === 'All' || spec.category === category) &&
        (needle === '' ||
          spec.name.includes(needle) ||
          t(spec.descKey).toUpperCase().includes(needle)),
    )
  }, [query, category, lang])

  const pick = (spec: FunctionSpec): void => {
    setPicked(spec)
    setFormula(`=${spec.name}(${spec.syntax.endsWith('()') ? ')' : ''}`)
    setError(null)
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog insert-function-dialog"
        role="dialog"
        aria-label="Insert Function"
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgFnTitle', { target: targetLabel })}</header>
        <div className="fn-filter-row">
          <input
            autoFocus
            placeholder={t('dlgFnSearchPlaceholder')}
            value={query}
            onChange={(event) => setQuery(event.target.value)}
          />
          <Dropdown
            value={category}
            options={CATEGORIES.map((name) => ({
              value: name,
              label: CATEGORY_LABELS[name] ? t(CATEGORY_LABELS[name]) : name,
            }))}
            onPick={setCategory}
          />
        </div>
        <div className="fn-list" role="listbox">
          {matches.map((spec) => (
            <button
              key={spec.name}
              className={`fn-row${picked?.name === spec.name ? ' active' : ''}`}
              role="option"
              aria-selected={picked?.name === spec.name}
              onClick={() => pick(spec)}
            >
              <strong>{spec.name}</strong>
              <span>{t(spec.descKey)}</span>
            </button>
          ))}
          {matches.length === 0 && <p className="dialog-note">{t('dlgFnNoMatch')}</p>}
        </div>
        {picked && (
          <p className="dialog-note fn-syntax">
            <code>{picked.syntax}</code>
          </p>
        )}
        <label className="fn-formula">
          {t('dlgFnFormula')}
          <input
            value={formula}
            placeholder={t('dlgFnFormulaPlaceholder')}
            onChange={(event) => setFormula(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault()
                const failure = onApply(formula)
                setError(failure)
                if (failure === null) onClose()
              }
            }}
          />
        </label>
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            disabled={formula.trim() === ''}
            onClick={() => {
              const failure = onApply(formula)
              setError(failure)
              if (failure === null) onClose()
            }}
          >
            {t('dlgFnInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}
