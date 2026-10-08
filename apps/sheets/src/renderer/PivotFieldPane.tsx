import { useEffect, useRef, useState } from 'react'

import { Dropdown } from '@genoffice/ui'
import type {
  PivotAggregation,
  PivotShowDataAs,
} from '@genoffice/xlsx-gateway/domain/pivot-value-modes'

import { useI18n, type StringKey, type TFunc } from './i18n/locale'
import { useModalDialog } from './modal-dialog'
import {
  PIVOT_AREAS,
  areaLength,
  isFieldUsed,
  moveChip,
  removeChip,
  setFieldGrouping,
  setFieldLabelFilter,
  setPageItem,
  toggleField,
  updateValueSpec,
  type PivotArea,
  type PivotChipRef,
  type PivotDragSource,
  type PivotField,
  type PivotGroupingOption,
  type PivotLabelFilterOption,
  type PivotLayoutModel,
  type PivotValueFilterOption,
  type PivotValueSpec,
} from './pivot-field-model'

/// Excel's PivotTable Fields task pane: the field list on top, the four drop
/// areas below. Chips drag between and within areas (pointer events) and
/// carry a ▾ menu with the move/remove/settings commands.

const AREA_LABELS: Record<PivotArea, StringKey> = {
  filters: 'dlgPivotAreaFilters',
  columns: 'dlgPivotAreaColumns',
  rows: 'dlgPivotAreaRows',
  values: 'dlgPivotAreaValues',
}

const MOVE_TO_LABELS: Record<PivotArea, StringKey> = {
  filters: 'dlgPivotMoveToFilters',
  columns: 'dlgPivotMoveToColumns',
  rows: 'dlgPivotMoveToRows',
  values: 'dlgPivotMoveToValues',
}

/// Every Excel "Summarize Values By" entry; the engine covers the first seven.
const AGG_OPTIONS: readonly { value: string; labelKey: StringKey; supported: boolean }[] = [
  { value: 'sum', labelKey: 'dlgPivotAggSum', supported: true },
  { value: 'count', labelKey: 'dlgPivotAggCount', supported: true },
  { value: 'average', labelKey: 'dlgPivotAggAverage', supported: true },
  { value: 'max', labelKey: 'dlgPivotAggMax', supported: true },
  { value: 'min', labelKey: 'dlgPivotAggMin', supported: true },
  { value: 'product', labelKey: 'dlgPivotAggProduct', supported: true },
  { value: 'countNums', labelKey: 'dlgPivotAggCountNums', supported: true },
  { value: 'stdDev', labelKey: 'dlgPivotAggStdDev', supported: false },
  { value: 'stdDevp', labelKey: 'dlgPivotAggStdDevp', supported: false },
  { value: 'var', labelKey: 'dlgPivotAggVar', supported: false },
  { value: 'varp', labelKey: 'dlgPivotAggVarp', supported: false },
]

/// Every Excel "Show Values As" entry; '' = No Calculation. Modes the engine
/// cannot bake/recompute are listed but not pickable.
const SHOW_AS_OPTIONS: readonly { value: string; labelKey: StringKey; supported: boolean }[] = [
  { value: '', labelKey: 'dlgPivotShowAsNormal', supported: true },
  { value: 'percentOfTotal', labelKey: 'dlgPivotShowAsPctTotal', supported: true },
  { value: 'percentOfCol', labelKey: 'dlgPivotShowAsPctCol', supported: true },
  { value: 'percentOfRow', labelKey: 'dlgPivotShowAsPctRow', supported: true },
  { value: 'percentOf', labelKey: 'dlgPivotShowAsPctOf', supported: false },
  { value: 'percentOfParentRow', labelKey: 'dlgPivotShowAsPctParentRow', supported: true },
  { value: 'percentOfParentCol', labelKey: 'dlgPivotShowAsPctParentCol', supported: true },
  { value: 'percentOfParent', labelKey: 'dlgPivotShowAsPctParent', supported: false },
  { value: 'difference', labelKey: 'dlgPivotShowAsDifference', supported: false },
  { value: 'percentDiff', labelKey: 'dlgPivotShowAsPctDifference', supported: false },
  { value: 'runTotal', labelKey: 'dlgPivotShowAsRunningTotal', supported: false },
  { value: 'percentOfRunningTotal', labelKey: 'dlgPivotShowAsPctRunningTotal', supported: false },
  { value: 'rankAscending', labelKey: 'dlgPivotShowAsRankAsc', supported: false },
  { value: 'rankDescending', labelKey: 'dlgPivotShowAsRankDesc', supported: false },
  { value: 'index', labelKey: 'dlgPivotShowAsIndex', supported: true },
]

/// Dropdown value for "(All)"; a NUL prefix cannot collide with a source label
/// (blank members are the empty string).
const ALL_ITEMS = '\u0000all'

function fieldLabel(t: TFunc, fields: readonly PivotField[], fieldIndex: number): string {
  return fields[fieldIndex]?.label ?? t('appFieldN', { n: fieldIndex + 1 })
}

function valueCaption(t: TFunc, fields: readonly PivotField[], value: PivotValueSpec): string {
  if (value.name?.trim()) return value.name
  if (value.formula !== undefined) return value.calcName?.trim() || t('dlgPivotCalcNamePlaceholder')
  const agg = AGG_OPTIONS.find((option) => option.value === value.agg)
  return t('dlgPivotValueCaption', {
    agg: agg ? t(agg.labelKey) : value.agg,
    field: fieldLabel(t, fields, value.fieldIndex),
  })
}

function chipLabel(
  t: TFunc,
  fields: readonly PivotField[],
  model: PivotLayoutModel,
  ref: PivotChipRef,
): string {
  switch (ref.area) {
    case 'filters':
      return fieldLabel(t, fields, model.filters[ref.index]?.fieldIndex ?? -1)
    case 'columns':
      return fieldLabel(t, fields, model.columns[ref.index] ?? -1)
    case 'rows':
      return fieldLabel(t, fields, model.rows[ref.index] ?? -1)
    case 'values': {
      const value = model.values[ref.index]
      return value ? valueCaption(t, fields, value) : ''
    }
  }
}

interface DragState {
  readonly source: PivotDragSource
  readonly label: string
  readonly startX: number
  readonly startY: number
  readonly x: number
  readonly y: number
  readonly active: boolean
  readonly over: { area: PivotArea; index: number } | null
}

interface MenuState {
  readonly ref: PivotChipRef
  readonly x: number
  readonly y: number
}

type SettingsTarget = PivotChipRef

export function PivotFieldPane({
  fields,
  model,
  onModelChange,
  deferred,
  onDeferredChange,
  dirty,
  error,
  onUpdate,
  onClose,
  onGetFieldMembers,
}: {
  readonly fields: readonly PivotField[]
  readonly model: PivotLayoutModel
  readonly onModelChange: (next: PivotLayoutModel) => void
  readonly deferred: boolean
  readonly onDeferredChange: (deferred: boolean) => void
  /// Deferred changes not yet applied.
  readonly dirty: boolean
  readonly error: string | null
  readonly onUpdate: () => void
  readonly onClose: () => void
  readonly onGetFieldMembers: (fieldIndex: number) => string[]
}): React.JSX.Element {
  const { t } = useI18n()
  const paneRef = useRef<HTMLDivElement>(null)
  const [drag, setDrag] = useState<DragState | null>(null)
  const [menu, setMenu] = useState<MenuState | null>(null)
  const [settings, setSettings] = useState<SettingsTarget | null>(null)
  const dragRef = useRef(drag)
  dragRef.current = drag
  const dragging = drag !== null

  const beginDrag = (event: React.PointerEvent, source: PivotDragSource, label: string): void => {
    if (event.button !== 0) return
    setMenu(null)
    setDrag({
      source,
      label,
      startX: event.clientX,
      startY: event.clientY,
      x: event.clientX,
      y: event.clientY,
      active: false,
      over: null,
    })
  }

  useEffect(() => {
    if (!dragging) return
    const dropTarget = (x: number, y: number): { area: PivotArea; index: number } | null => {
      const element = document.elementFromPoint(x, y)
      const areaElement = element?.closest<HTMLElement>('[data-pivot-area]')
      const area = areaElement?.dataset['pivotArea'] as PivotArea | undefined
      if (!areaElement || !area || !PIVOT_AREAS.includes(area)) return null
      const chips = [...areaElement.querySelectorAll<HTMLElement>('[data-chip-index]')]
      for (const chip of chips) {
        const rect = chip.getBoundingClientRect()
        if (y < rect.top + rect.height / 2)
          return { area, index: Number(chip.dataset['chipIndex']) }
      }
      return { area, index: chips.length }
    }
    const onMove = (event: PointerEvent): void => {
      const current = dragRef.current
      if (!current) return
      const active =
        current.active ||
        Math.abs(event.clientX - current.startX) > 4 ||
        Math.abs(event.clientY - current.startY) > 4
      setDrag({
        ...current,
        x: event.clientX,
        y: event.clientY,
        active,
        over: active ? dropTarget(event.clientX, event.clientY) : null,
      })
    }
    const onUp = (): void => {
      const current = dragRef.current
      setDrag(null)
      if (!current?.active || !current.over) return
      onModelChange(moveChip(model, fields, current.source, current.over))
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
    window.addEventListener('pointercancel', onUp)
    return () => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
      window.removeEventListener('pointercancel', onUp)
    }
  }, [dragging, model, fields, onModelChange])

  useEffect(() => {
    if (!menu) return
    const close = (event: Event): void => {
      if ((event.target as HTMLElement | null)?.closest('.pivot-chip-menu')) return
      setMenu(null)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setMenu(null)
    }
    document.addEventListener('pointerdown', close, true)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('pointerdown', close, true)
      document.removeEventListener('keydown', onKey)
    }
  }, [menu])

  const runMenu = (ref: PivotChipRef, action: () => void): void => {
    setMenu(null)
    action()
  }

  const menuItems = (
    ref: PivotChipRef,
  ): { label: string; disabled?: boolean; onPick: () => void }[] => {
    const length = areaLength(model, ref.area)
    const move = (index: number) => () =>
      onModelChange(moveChip(model, fields, ref, { area: ref.area, index }))
    const fieldIndex =
      ref.area === 'values'
        ? (model.values[ref.index]?.fieldIndex ?? -1)
        : ref.area === 'filters'
          ? (model.filters[ref.index]?.fieldIndex ?? -1)
          : ref.area === 'rows'
            ? (model.rows[ref.index] ?? -1)
            : (model.columns[ref.index] ?? -1)
    const items: { label: string; disabled?: boolean; onPick: () => void }[] = [
      { label: t('dlgPivotMoveUp'), disabled: ref.index === 0, onPick: move(ref.index - 1) },
      {
        label: t('dlgPivotMoveDown'),
        disabled: ref.index >= length - 1,
        onPick: move(ref.index + 2),
      },
      { label: t('dlgPivotMoveToBeginning'), disabled: ref.index === 0, onPick: move(0) },
      {
        label: t('dlgPivotMoveToEnd'),
        disabled: ref.index >= length - 1,
        onPick: move(length),
      },
    ]
    for (const area of PIVOT_AREAS) {
      if (area === ref.area) continue
      // Calculated fields only ever live in Values.
      if (fieldIndex < 0 && area !== 'values') continue
      items.push({
        label: t(MOVE_TO_LABELS[area]),
        onPick: () =>
          onModelChange(moveChip(model, fields, ref, { area, index: areaLength(model, area) })),
      })
    }
    items.push({
      label: t('dlgPivotRemoveField'),
      onPick: () => onModelChange(removeChip(model, ref)),
    })
    items.push({
      label: ref.area === 'values' ? t('dlgPivotValueFieldSettings') : t('dlgPivotFieldSettings'),
      onPick: () => setSettings(ref),
    })
    return items
  }

  const renderChip = (ref: PivotChipRef): React.JSX.Element => {
    const label = chipLabel(t, fields, model, ref)
    const isDragSource =
      drag?.active === true &&
      drag.source.area === ref.area &&
      (drag.source as PivotChipRef).index === ref.index
    const page = ref.area === 'filters' ? model.filters[ref.index] : undefined
    return (
      <div
        key={`${ref.area}-${ref.index}`}
        className={`pivot-chip${isDragSource ? ' is-dragging' : ''}`}
        data-chip-index={ref.index}
        onPointerDown={(event) => {
          if ((event.target as HTMLElement).closest('button')) return
          beginDrag(event, ref, label)
        }}
      >
        <span className="pivot-chip-label" title={label}>
          {label}
          {page !== undefined && (
            <small>{page.item === null ? t('dlgPivotAllItems') : page.item}</small>
          )}
        </span>
        <button
          type="button"
          aria-label={t('dlgPivotFieldMenu', { name: label })}
          aria-haspopup="menu"
          onClick={(event) => {
            const rect = event.currentTarget.getBoundingClientRect()
            setMenu({ ref, x: rect.right, y: rect.bottom })
          }}
        >
          ▾
        </button>
      </div>
    )
  }

  const renderArea = (area: PivotArea): React.JSX.Element => {
    const length = areaLength(model, area)
    const isOver = drag?.active === true && drag.over?.area === area
    const chips: React.JSX.Element[] = []
    for (let index = 0; index < length; index += 1) {
      if (isOver && drag.over?.index === index) {
        chips.push(<div key={`marker-${index}`} className="pivot-drop-marker" />)
      }
      chips.push(renderChip({ area, index }))
    }
    if (isOver && drag.over?.index === length) {
      chips.push(<div key="marker-end" className="pivot-drop-marker" />)
    }
    return (
      <section
        key={area}
        className={`pivot-area${isOver ? ' is-drop-target' : ''}`}
        data-pivot-area={area}
      >
        <strong>{t(AREA_LABELS[area])}</strong>
        <div className="pivot-area-chips">{chips}</div>
      </section>
    )
  }

  return (
    <div className="pivot-fields-pane" ref={paneRef} role="complementary">
      <header>
        <span>{t('dlgPivotPaneTitle')}</span>
        <button type="button" data-tip={t('appClose')} aria-label={t('appClose')} onClick={onClose}>
          ✕
        </button>
      </header>
      <div className="pivot-field-list">
        <strong>{t('dlgPivotPaneFieldList')}</strong>
        <ul>
          {fields.map((field, fieldIndex) => (
            <li
              key={fieldIndex}
              onPointerDown={(event) => {
                if ((event.target as HTMLElement).tagName === 'INPUT') return
                beginDrag(event, { area: 'list', fieldIndex }, field.label)
              }}
            >
              <label>
                <input
                  type="checkbox"
                  checked={isFieldUsed(model, fieldIndex)}
                  onChange={() => onModelChange(toggleField(model, fields, fieldIndex))}
                />
                <span title={field.label}>{field.label}</span>
              </label>
            </li>
          ))}
        </ul>
      </div>
      <div className="pivot-areas">
        <strong>{t('dlgPivotPaneDragHint')}</strong>
        <div className="pivot-area-grid">{PIVOT_AREAS.map(renderArea)}</div>
      </div>
      {error && (
        <p className="dialog-error" role="alert">
          {error}
        </p>
      )}
      <footer>
        <label>
          <input
            type="checkbox"
            checked={deferred}
            onChange={(event) => onDeferredChange(event.target.checked)}
          />
          {t('dlgPivotDeferUpdate')}
        </label>
        <button type="button" className="primary-action" disabled={!dirty} onClick={onUpdate}>
          {t('dlgPivotUpdate')}
        </button>
      </footer>
      {drag?.active && (
        <div className="pivot-drag-ghost" style={{ left: drag.x + 8, top: drag.y + 8 }}>
          {drag.label}
        </div>
      )}
      {menu && (
        <div className="pivot-chip-menu" role="menu" style={{ left: menu.x, top: menu.y }}>
          {menuItems(menu.ref).map((item, index) => (
            <button
              key={index}
              type="button"
              role="menuitem"
              disabled={item.disabled}
              onClick={() => runMenu(menu.ref, item.onPick)}
            >
              {item.label}
            </button>
          ))}
        </div>
      )}
      {settings &&
        (settings.area === 'values' ? (
          <ValueFieldSettingsDialog
            fields={fields}
            value={model.values[settings.index]!}
            onApply={(patch) => onModelChange(updateValueSpec(model, settings.index, patch))}
            onClose={() => setSettings(null)}
          />
        ) : (
          <FieldSettingsDialog
            fields={fields}
            model={model}
            target={settings}
            members={
              settings.area === 'filters'
                ? onGetFieldMembers(model.filters[settings.index]?.fieldIndex ?? -1)
                : []
            }
            onModelChange={onModelChange}
            onClose={() => setSettings(null)}
          />
        ))}
    </div>
  )
}

function ValueFieldSettingsDialog({
  fields,
  value,
  onApply,
  onClose,
}: {
  readonly fields: readonly PivotField[]
  readonly value: PivotValueSpec
  readonly onApply: (patch: Partial<PivotValueSpec>) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [draft, setDraft] = useState<PivotValueSpec>(value)
  const modal = useModalDialog(onClose)
  const isCalc = draft.formula !== undefined
  const sourceName = isCalc ? (draft.calcName ?? '') : fieldLabel(t, fields, draft.fieldIndex)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog pivot-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgPivotValueFieldSettingsTitle', { name: sourceName })}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgPivotValueFieldSettingsTitle', { name: sourceName })}</header>
        <div className="dialog-grid">
          {isCalc ? (
            <>
              <label>
                {t('dlgPivotCalcNamePlaceholder')}
                <input
                  type="text"
                  className="cell-input"
                  value={draft.calcName ?? ''}
                  onChange={(event) => setDraft({ ...draft, calcName: event.target.value })}
                />
              </label>
              <label>
                {t('dlgPivotCalcFormulaPlaceholder')}
                <input
                  type="text"
                  className="cell-input"
                  value={draft.formula ?? ''}
                  onChange={(event) => setDraft({ ...draft, formula: event.target.value })}
                />
              </label>
            </>
          ) : (
            <p className="dialog-note dialog-span">
              {t('dlgPivotSourceName', { name: sourceName })}
            </p>
          )}
          <label className="dialog-span">
            {t('dlgPivotCustomName')}
            <input
              type="text"
              className="cell-input"
              placeholder={valueCaption(t, fields, { ...draft, name: undefined })}
              value={draft.name ?? ''}
              onChange={(event) =>
                setDraft({
                  ...draft,
                  name: event.target.value === '' ? undefined : event.target.value,
                })
              }
            />
          </label>
          {!isCalc && (
            <label>
              {t('dlgPivotSummarizeBy')}
              <Dropdown
                value={draft.agg}
                options={AGG_OPTIONS.map((option) => ({
                  value: option.value,
                  label: option.supported
                    ? t(option.labelKey)
                    : `${t(option.labelKey)} — ${t('dlgPivotNotSupportedYet')}`,
                  disabled: !option.supported,
                }))}
                onPick={(agg) => setDraft({ ...draft, agg: agg as PivotAggregation })}
              />
            </label>
          )}
          <label>
            {t('dlgPivotShowAs')}
            <Dropdown
              value={draft.showDataAs ?? ''}
              options={SHOW_AS_OPTIONS.map((option) => ({
                value: option.value,
                label: option.supported
                  ? t(option.labelKey)
                  : `${t(option.labelKey)} — ${t('dlgPivotNotSupportedYet')}`,
                disabled: !option.supported,
              }))}
              onPick={(mode) =>
                setDraft({
                  ...draft,
                  showDataAs: mode === '' ? undefined : (mode as PivotShowDataAs),
                })
              }
            />
          </label>
          <ValueFilterFields
            filter={draft.filter}
            onChange={(filter) => setDraft({ ...draft, filter })}
          />
        </div>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            onClick={() => {
              onApply({
                agg: draft.agg,
                showDataAs: draft.showDataAs,
                name: draft.name?.trim() || undefined,
                calcName: draft.calcName,
                formula: draft.formula,
                filter: draft.filter,
              })
              onClose()
            }}
          >
            {t('dlgApply')}
          </button>
        </div>
      </div>
    </div>
  )
}

function ValueFilterFields({
  filter,
  onChange,
}: {
  readonly filter: PivotValueFilterOption | undefined
  readonly onChange: (filter: PivotValueFilterOption | undefined) => void
}): React.JSX.Element {
  const { t } = useI18n()
  return (
    <>
      <label>
        {t('dlgPivotValueFilter')}
        <Dropdown
          value={filter?.op ?? ''}
          options={[
            { value: '', label: t('dlgPivotFilterNone') },
            { value: 'top', label: t('dlgPivotFilterTopN') },
            { value: 'greaterThan', label: t('dlgPivotFilterGreaterThan') },
            { value: 'between', label: t('dlgPivotFilterBetween') },
          ]}
          onPick={(picked) => {
            const op = picked as PivotValueFilterOption['op'] | ''
            if (op === '') onChange(undefined)
            else if (op === 'top') onChange({ op, count: filter?.count ?? 10 })
            else if (op === 'greaterThan') onChange({ op, from: filter?.from ?? 0 })
            else onChange({ op, from: filter?.from ?? 0, to: filter?.to ?? 100 })
          }}
        />
      </label>
      {filter?.op === 'top' && (
        <label>
          {t('dlgPivotTopNCount')}
          <input
            type="number"
            className="cell-input"
            min={1}
            value={filter.count ?? 10}
            onChange={(event) => onChange({ op: 'top', count: Number(event.target.value) })}
          />
        </label>
      )}
      {filter?.op === 'greaterThan' && (
        <label>
          {t('dlgPivotGreaterThanValue')}
          <input
            type="number"
            className="cell-input"
            value={filter.from ?? 0}
            onChange={(event) => onChange({ op: 'greaterThan', from: Number(event.target.value) })}
          />
        </label>
      )}
      {filter?.op === 'between' && (
        <>
          <label>
            {t('dlgPivotBetweenFrom')}
            <input
              type="number"
              className="cell-input"
              value={filter.from ?? 0}
              onChange={(event) =>
                onChange({ op: 'between', from: Number(event.target.value), to: filter.to ?? 100 })
              }
            />
          </label>
          <label>
            {t('dlgPivotBetweenTo')}
            <input
              type="number"
              className="cell-input"
              value={filter.to ?? 100}
              onChange={(event) =>
                onChange({ op: 'between', from: filter.from ?? 0, to: Number(event.target.value) })
              }
            />
          </label>
        </>
      )}
    </>
  )
}

/// Row/column fields: grouping and label filter. Report filters: the selected
/// item.
function FieldSettingsDialog({
  fields,
  model,
  target,
  members,
  onModelChange,
  onClose,
}: {
  readonly fields: readonly PivotField[]
  readonly model: PivotLayoutModel
  readonly target: PivotChipRef
  readonly members: readonly string[]
  readonly onModelChange: (next: PivotLayoutModel) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const modal = useModalDialog(onClose)
  const fieldIndex =
    target.area === 'filters'
      ? (model.filters[target.index]?.fieldIndex ?? -1)
      : target.area === 'rows'
        ? (model.rows[target.index] ?? -1)
        : (model.columns[target.index] ?? -1)
  const name = fieldLabel(t, fields, fieldIndex)
  const [grouping, setGrouping] = useState<PivotGroupingOption | null>(
    model.groupings[fieldIndex] ?? null,
  )
  const [labelFilter, setLabelFilter] = useState<PivotLabelFilterOption | null>(
    model.labelFilters[fieldIndex] ?? null,
  )
  const [item, setItem] = useState<string | null>(model.filters[target.index]?.item ?? null)
  const groupingValue =
    grouping === null ? '' : grouping.kind === 'date' ? `date:${grouping.dateUnit}` : 'range'
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog pivot-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgPivotFieldSettingsTitle', { name })}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgPivotFieldSettingsTitle', { name })}</header>
        <div className="dialog-grid">
          {target.area === 'filters' ? (
            <label className="dialog-span">
              {t('dlgPivotReportFilterItem')}
              <Dropdown
                value={item ?? ALL_ITEMS}
                options={[
                  { value: ALL_ITEMS, label: t('dlgPivotAllItems') },
                  ...members.map((member) => ({
                    value: member,
                    label: member === '' ? t('appBlank') : member,
                  })),
                ]}
                onPick={(picked) => setItem(picked === ALL_ITEMS ? null : picked)}
              />
            </label>
          ) : (
            <>
              <label>
                {t('dlgPivotGrouping')}
                <Dropdown
                  value={groupingValue}
                  options={[
                    { value: '', label: t('dlgPivotGroupNone') },
                    { value: 'date:year', label: t('dlgPivotGroupYear') },
                    { value: 'date:quarter', label: t('dlgPivotGroupQuarter') },
                    { value: 'date:month', label: t('dlgPivotGroupMonth') },
                    { value: 'range', label: t('dlgPivotGroupRange') },
                  ]}
                  onPick={(picked) => {
                    if (picked === '') setGrouping(null)
                    else if (picked === 'range')
                      setGrouping({
                        kind: 'range',
                        rangeStep: grouping?.kind === 'range' ? grouping.rangeStep : 100,
                      })
                    else
                      setGrouping({
                        kind: 'date',
                        dateUnit: picked.slice('date:'.length) as 'year' | 'quarter' | 'month',
                      })
                  }}
                />
              </label>
              {grouping?.kind === 'range' && (
                <label>
                  {t('dlgPivotRangeStep')}
                  <input
                    type="number"
                    className="cell-input"
                    min={1}
                    value={grouping.rangeStep}
                    onChange={(event) =>
                      setGrouping({ kind: 'range', rangeStep: Number(event.target.value) })
                    }
                  />
                </label>
              )}
              <label>
                {t('dlgPivotLabelFilter')}
                <Dropdown
                  value={labelFilter?.op ?? ''}
                  options={[
                    { value: '', label: t('dlgPivotFilterNone') },
                    { value: 'equal', label: t('dlgPivotFilterEqual') },
                    { value: 'contains', label: t('dlgPivotFilterContains') },
                    { value: 'beginsWith', label: t('dlgPivotFilterBeginsWith') },
                  ]}
                  onPick={(picked) => {
                    const op = picked as PivotLabelFilterOption['op'] | ''
                    setLabelFilter(op === '' ? null : { op, value: labelFilter?.value ?? '' })
                  }}
                />
              </label>
              {labelFilter && (
                <label>
                  {t('dlgPivotLabelFilterText')}
                  <input
                    type="text"
                    className="cell-input"
                    placeholder={t('dlgPivotFilterTextPlaceholder')}
                    value={labelFilter.value}
                    onChange={(event) =>
                      setLabelFilter({ op: labelFilter.op, value: event.target.value })
                    }
                  />
                </label>
              )}
            </>
          )}
        </div>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            onClick={() => {
              if (target.area === 'filters') {
                onModelChange(setPageItem(model, target.index, item))
              } else {
                onModelChange(
                  setFieldLabelFilter(
                    setFieldGrouping(model, fieldIndex, grouping),
                    fieldIndex,
                    labelFilter,
                  ),
                )
              }
              onClose()
            }}
          >
            {t('dlgApply')}
          </button>
        </div>
      </div>
    </div>
  )
}
