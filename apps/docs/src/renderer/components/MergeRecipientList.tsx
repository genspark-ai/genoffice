import { useRef, useState } from 'react'
import type { MergeDataSource } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import {
  type FilterCondition,
  type FilterJoin,
  type FilterOperator,
  type RecipientRow,
  type SortKey,
} from '../mail-merge/recipients'

/**
 * The recipient list: the records the merge draws on, and the ones the user
 * takes out of it.
 *
 * The Filter and Sort buttons open a two-tab dialog (Filter Records / Sort
 * Records). Filter rows are field + operator + value, joined with and/or;
 * Sort offers up to three levels. Changes are drafted — OK applies them,
 * Cancel leaves the list as it was.
 */

export interface MergeRecipientListProps {
  /** the data source the records come from, null when there is none yet */
  source: MergeDataSource | null
  /** the records to show: source order, or the user's filter and sort */
  rows: readonly RecipientRow[]
  /** the indexes of the records the user took out of the merge */
  excluded: ReadonlySet<number>
  /** take a record out of the merge, or put it back */
  onExcluded: (excluded: Set<number>) => void
  /** the filter conditions the list applies */
  conditions: readonly FilterCondition[]
  onConditions: (conditions: FilterCondition[]) => void
  /** the sort levels the list applies, at most three */
  sorts: readonly SortKey[]
  onSorts: (sorts: SortKey[]) => void
  onClose: () => void
}

const OPERATORS: readonly FilterOperator[] = [
  'eq',
  'ne',
  'gt',
  'gte',
  'lt',
  'lte',
  'contains',
  'notContains',
  'blank',
  'notBlank',
]
/** Word shows six condition rows and grows the list as they fill up */
const MIN_FILTER_ROWS = 6
const MAX_FILTER_ROWS = 20
const SORT_LEVELS = 3

const emptyCondition = (join: FilterJoin = 'and'): FilterCondition => ({
  field: '',
  operator: 'eq',
  value: '',
  join,
})

const emptySort: SortKey = { field: '', direction: 'asc' }

export function MergeRecipientList({
  source,
  rows,
  excluded,
  onExcluded,
  conditions,
  onConditions,
  sorts,
  onSorts,
  onClose,
}: MergeRecipientListProps) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const fields = source?.fields ?? []
  const total = source?.records.length ?? 0
  const kept = rows.filter((row) => !excluded.has(row.index)).length
  const [tab, setTab] = useState<'filter' | 'sort' | null>(null)
  const [draftConditions, setDraftConditions] = useState<FilterCondition[]>([])
  const [draftSorts, setDraftSorts] = useState<SortKey[]>([])

  // opening the dialog copies what the list applies into the draft it edits;
  // switching tabs keeps that draft
  const openTab = (next: 'filter' | 'sort') => {
    if (!tab) {
      setDraftConditions(
        conditions.length > 0 ? conditions.map((c) => ({ ...c })) : [emptyCondition()],
      )
      setDraftSorts(
        Array.from({ length: SORT_LEVELS }, (_, i) =>
          sorts[i] ? { ...sorts[i] } : { ...emptySort },
        ),
      )
    }
    setTab(next)
  }

  const operatorLabel = (operator: FilterOperator): string => {
    switch (operator) {
      case 'eq':
        return t('ribbonMailingsOpEquals')
      case 'ne':
        return t('ribbonMailingsOpNotEquals')
      case 'gt':
        return t('ribbonMailingsOpGreater')
      case 'gte':
        return t('ribbonMailingsOpGreaterOrEqual')
      case 'lt':
        return t('ribbonMailingsOpLess')
      case 'lte':
        return t('ribbonMailingsOpLessOrEqual')
      case 'contains':
        return t('ribbonMailingsOpContains')
      case 'notContains':
        return t('ribbonMailingsOpNotContains')
      case 'blank':
        return t('ribbonMailingsOpBlank')
      case 'notBlank':
        return t('ribbonMailingsOpNotBlank')
    }
  }

  const updateCondition = (index: number, patch: Partial<FilterCondition>) => {
    const next = draftConditions.map((c) => ({ ...c }))
    while (next.length <= index) next.push(emptyCondition())
    next[index] = { ...next[index], ...patch }
    setDraftConditions(next)
  }

  const updateSort = (index: number, patch: Partial<SortKey>) => {
    const next = draftSorts.map((s) => ({ ...s }))
    while (next.length <= index) next.push({ ...emptySort })
    next[index] = { ...next[index], ...patch }
    setDraftSorts(next)
  }

  const apply = () => {
    onConditions(draftConditions.filter((c) => c.field !== ''))
    onSorts(draftSorts.filter((s) => s.field !== '').slice(0, SORT_LEVELS))
    setTab(null)
  }

  const activeConditions = conditions.filter((c) => c.field !== '').length
  const firstSort = sorts.find((s) => s.field !== '')

  const allKept = rows.length > 0 && kept === rows.length
  const partlyKept = kept > 0 && !allKept

  const onChangeAll = (on: boolean) => {
    const next = new Set(excluded)
    for (const row of rows) {
      if (on) next.delete(row.index)
      else next.add(row.index)
    }
    onExcluded(next)
  }

  const filterRows = Math.min(
    Math.max(MIN_FILTER_ROWS, draftConditions.length + 1),
    MAX_FILTER_ROWS,
  )

  // the dialog behaves like a window: drag its title to move it, drag the
  // corner to resize it; the list grows into the extra height
  const modalRef = useRef<HTMLDivElement>(null)
  const [offset, setOffset] = useState({ x: 0, y: 0 })
  const [size, setSize] = useState<{ width: number; height: number } | null>(null)
  const moveRef = useRef<{
    x: number
    y: number
    offset: { x: number; y: number }
    rect: DOMRect
  } | null>(null)
  const resizeRef = useRef<{ x: number; y: number; width: number; height: number } | null>(null)

  const onTitlePointerDown = (e: React.PointerEvent<HTMLHeadingElement>) => {
    if (e.button !== 0) return
    const rect = modalRef.current?.getBoundingClientRect()
    if (!rect) return
    e.currentTarget.setPointerCapture(e.pointerId)
    moveRef.current = { x: e.clientX, y: e.clientY, offset, rect }
  }

  const onTitlePointerMove = (e: React.PointerEvent<HTMLHeadingElement>) => {
    const start = moveRef.current
    if (!start) return
    // keep a strip of the window on screen so it cannot be dragged away
    const left = Math.min(
      Math.max(start.rect.left + start.offset.x + e.clientX - start.x, 80 - start.rect.width),
      window.innerWidth - 80,
    )
    const top = Math.min(
      Math.max(start.rect.top + start.offset.y + e.clientY - start.y, 0),
      window.innerHeight - 48,
    )
    setOffset({ x: left - start.rect.left, y: top - start.rect.top })
  }

  const endMove = (e: React.PointerEvent<HTMLHeadingElement>) => {
    moveRef.current = null
    if (e.currentTarget.hasPointerCapture(e.pointerId))
      e.currentTarget.releasePointerCapture(e.pointerId)
  }

  const onResizePointerDown = (e: React.PointerEvent<HTMLSpanElement>) => {
    if (e.button !== 0) return
    const rect = modalRef.current?.getBoundingClientRect()
    if (!rect) return
    e.currentTarget.setPointerCapture(e.pointerId)
    resizeRef.current = { x: e.clientX, y: e.clientY, width: rect.width, height: rect.height }
  }

  const onResizePointerMove = (e: React.PointerEvent<HTMLSpanElement>) => {
    const start = resizeRef.current
    if (!start) return
    setSize({
      width: Math.min(Math.max(start.width + e.clientX - start.x, 520), window.innerWidth - 16),
      height: Math.min(Math.max(start.height + e.clientY - start.y, 320), window.innerHeight - 16),
    })
  }

  const endResize = (e: React.PointerEvent<HTMLSpanElement>) => {
    resizeRef.current = null
    if (e.currentTarget.hasPointerCapture(e.pointerId))
      e.currentTarget.releasePointerCapture(e.pointerId)
  }

  // Esc closes the Filter/Sort dialog before it closes the recipient list
  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape' && tab) {
      e.preventDefault()
      e.stopPropagation()
      setTab(null)
      return
    }
    modalKeys.onKeyDown(e)
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className={
          size
            ? 'modal modal-mailmerge modal-recipients mailmerge-resizable'
            : 'modal modal-mailmerge modal-recipients'
        }
        ref={modalRef}
        style={{
          transform: `translate(${offset.x}px, ${offset.y}px)`,
          ...(size ? { width: size.width, height: size.height } : null),
        }}
      >
        <h2
          className="mailmerge-drag-title"
          onPointerDown={onTitlePointerDown}
          onPointerMove={onTitlePointerMove}
          onPointerUp={endMove}
          onPointerCancel={endMove}
        >
          {t('ribbonMailingsRecipientList')}
        </h2>

        <div className="mailmerge-source-line">
          <span className="mailmerge-label">{t('ribbonMailingsSource')}</span>
          <span className="mailmerge-source-name" title={source?.path}>
            {source ? source.name : t('ribbonMailingsNoSource')}
          </span>
          {source?.sheet && (
            <span className="mailmerge-sheet-name">
              {t('ribbonMailingsSheetLabel', { sheet: source.sheet })}
            </span>
          )}
        </div>

        {total > 0 ? (
          <div className="mailmerge-section mailmerge-list-section">
            <div className="mailmerge-toolbar">
              <button
                type="button"
                className={activeConditions > 0 ? 'mailmerge-toggle-on' : undefined}
                onClick={() => openTab('filter')}
              >
                {t('ribbonMailingsFilter')}
              </button>
              <button
                type="button"
                className={firstSort ? 'mailmerge-toggle-on' : undefined}
                onClick={() => openTab('sort')}
              >
                {t('ribbonMailingsSort')}
              </button>
            </div>

            <div className="mailmerge-recipients">
              <table className="mailmerge-recipient-table">
                <thead>
                  <tr>
                    <th className="mailmerge-recipient-check">
                      <input
                        type="checkbox"
                        aria-label={t('ribbonMailingsSelectAll')}
                        checked={allKept}
                        ref={(el) => {
                          if (el) el.indeterminate = partlyKept
                        }}
                        onChange={(e) => onChangeAll(e.target.checked)}
                      />
                    </th>
                    {fields.map((name) => (
                      <th key={name}>
                        {name}
                        {firstSort?.field === name && (
                          <span className="mailmerge-sort-arrow">
                            {firstSort.direction === 'asc' ? '▲' : '▼'}
                          </span>
                        )}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map(({ record, index }) => (
                    <tr
                      key={index}
                      className={excluded.has(index) ? 'mailmerge-recipient-off' : undefined}
                    >
                      <td className="mailmerge-recipient-check">
                        <input
                          type="checkbox"
                          checked={!excluded.has(index)}
                          onChange={() => {
                            const next = new Set(excluded)
                            if (next.has(index)) next.delete(index)
                            else next.add(index)
                            onExcluded(next)
                          }}
                        />
                      </td>
                      {fields.map((name) => (
                        <td key={name}>{record[name] ?? ''}</td>
                      ))}
                    </tr>
                  ))}
                  {rows.length === 0 && (
                    <tr>
                      <td className="mailmerge-no-results" colSpan={fields.length + 1}>
                        {t('ribbonMailingsNoFilterResults')}
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <div className="mailmerge-count">
              {t('ribbonMailingsRecipientCount', { n: kept, total })}
            </div>
          </div>
        ) : (
          <div className="mailmerge-section">
            <div className="mailmerge-empty">{t('ribbonMailingsNoRecipients')}</div>
          </div>
        )}

        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonMailingsClose')}</button>
        </div>
        <span
          className="mailmerge-resize-handle"
          aria-hidden="true"
          onPointerDown={onResizePointerDown}
          onPointerMove={onResizePointerMove}
          onPointerUp={endResize}
          onPointerCancel={endResize}
        />
      </div>

      {tab && (
        <div
          className="modal-backdrop mailmerge-cond-backdrop"
          onMouseDown={(e) => e.target === e.currentTarget && setTab(null)}
        >
          <div className="modal mailmerge-cond-modal">
            <h2>{t('ribbonMailingsFilterSortDialog')}</h2>

            <div className="mailmerge-tabs" role="tablist">
              {(['filter', 'sort'] as const).map((name) => (
                <button
                  key={name}
                  type="button"
                  role="tab"
                  aria-selected={tab === name}
                  className={tab === name ? 'mailmerge-tab mailmerge-tab-on' : 'mailmerge-tab'}
                  onClick={() => setTab(name)}
                >
                  {name === 'filter' ? t('ribbonMailingsTabFilter') : t('ribbonMailingsTabSort')}
                </button>
              ))}
            </div>

            {tab === 'filter' ? (
              <div className="mailmerge-cond-scroll">
                <div className="mailmerge-cond-head">
                  <span />
                  <span>{t('ribbonMailingsFieldColumn')}</span>
                  <span>{t('ribbonMailingsOperatorColumn')}</span>
                  <span>{t('ribbonMailingsValueColumn')}</span>
                </div>
                {Array.from({ length: filterRows }, (_, index) => {
                  const condition = draftConditions[index] ?? emptyCondition()
                  return (
                    <div className="mailmerge-cond-row" key={index}>
                      {index === 0 ? (
                        <span className="mailmerge-join-blank" />
                      ) : (
                        <select
                          className="mailmerge-join"
                          aria-label={t('ribbonMailingsCondAnd')}
                          value={condition.join}
                          onChange={(e) =>
                            updateCondition(index, { join: e.target.value as FilterJoin })
                          }
                        >
                          <option value="and">{t('ribbonMailingsCondAnd')}</option>
                          <option value="or">{t('ribbonMailingsCondOr')}</option>
                        </select>
                      )}
                      <select
                        className="mailmerge-cond-field"
                        value={condition.field}
                        onChange={(e) => updateCondition(index, { field: e.target.value })}
                      >
                        <option value="">{t('ribbonMailingsFilterAllFields')}</option>
                        {fields.map((name) => (
                          <option key={name} value={name}>
                            {name}
                          </option>
                        ))}
                      </select>
                      <select
                        className="mailmerge-cond-op"
                        value={condition.operator}
                        onChange={(e) =>
                          updateCondition(index, { operator: e.target.value as FilterOperator })
                        }
                      >
                        {OPERATORS.map((operator) => (
                          <option key={operator} value={operator}>
                            {operatorLabel(operator)}
                          </option>
                        ))}
                      </select>
                      <input
                        className="mailmerge-cond-value"
                        value={condition.value}
                        disabled={
                          condition.operator === 'blank' || condition.operator === 'notBlank'
                        }
                        placeholder={t('ribbonMailingsFilterPlaceholder')}
                        onChange={(e) => updateCondition(index, { value: e.target.value })}
                      />
                    </div>
                  )
                })}
              </div>
            ) : (
              <div className="mailmerge-sort-levels">
                {Array.from({ length: SORT_LEVELS }, (_, index) => {
                  const level = draftSorts[index] ?? emptySort
                  const label =
                    index === 0
                      ? t('ribbonMailingsSortBy')
                      : index === 1
                        ? t('ribbonMailingsSortThenBy')
                        : t('ribbonMailingsSortThenByThird')
                  return (
                    <div className="mailmerge-sort-row" key={index}>
                      <span className="mailmerge-sort-label">{label}</span>
                      <select
                        className="mailmerge-cond-field"
                        value={level.field}
                        onChange={(e) => updateSort(index, { field: e.target.value })}
                      >
                        <option value="">{t('ribbonMailingsSortNone')}</option>
                        {fields.map((name) => (
                          <option key={name} value={name}>
                            {name}
                          </option>
                        ))}
                      </select>
                      <label className="mailmerge-radio">
                        <input
                          type="radio"
                          name={`merge-sort-${index}`}
                          disabled={!level.field}
                          checked={level.direction === 'asc'}
                          onChange={() => updateSort(index, { direction: 'asc' })}
                        />
                        {t('ribbonMailingsSortAsc')}
                      </label>
                      <label className="mailmerge-radio">
                        <input
                          type="radio"
                          name={`merge-sort-${index}`}
                          disabled={!level.field}
                          checked={level.direction === 'desc'}
                          onChange={() => updateSort(index, { direction: 'desc' })}
                        />
                        {t('ribbonMailingsSortDesc')}
                      </label>
                    </div>
                  )
                })}
              </div>
            )}

            <div className="modal-actions mailmerge-cond-footer">
              <button
                type="button"
                onClick={() => {
                  if (tab === 'filter') setDraftConditions([emptyCondition()])
                  else setDraftSorts(Array.from({ length: SORT_LEVELS }, () => ({ ...emptySort })))
                }}
              >
                {t('ribbonMailingsClearAll')}
              </button>
              <span className="mailmerge-cond-spacer" />
              <button type="button" onClick={() => setTab(null)}>
                {t('ribbonCancel')}
              </button>
              <button type="button" className="primary" onClick={apply}>
                {t('ribbonOk')}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
