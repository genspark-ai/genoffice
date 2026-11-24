import { createIdentifier, ICommandService, LocaleService, numfmt } from '@univerjs/core'
import type { IFilterColumn } from '@univerjs/preset-sheets-filter'
import { FilterBy, type FilterModel } from '@univerjs/sheets-filter'
import { SheetsUIPart } from '@univerjs/sheets-ui'
import { ComponentContainer, ComponentManager, useComponentsOfPart } from '@univerjs/ui'
import { type ComponentType, useEffect, useRef, useState } from 'react'
import {
  buildFilterValueTree,
  filterModelFor,
  type FilterPage,
  type FilterValueSource,
  type FilterValueTreeInput,
  isLargeFilterSheet,
  readLargeFilterPage,
} from './large-filter'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'
import { largeFilterSelectValuesMessage } from './i18n/large-filter'
import { t as appText } from './i18n/locale'

const PANEL_KEY = 'FILTER_PANEL_POPUP'
const VALUES_SERVICE = 'sheets-filter.generate-filter-values.service'
const CHANGE_BY = 'sheet.operation.apply-filter'
const CLOSE_PANEL = 'sheet.operation.close-filter-panel'
const PAGE_SIZE = 200
// Excel's checklist limits the number of distinct items it offers at once.
const MAX_BULK_VALUES = 10_000

interface Target {
  state: LazyWorkbookState
  model: FilterModel
  column: number
  apply: (criteria: IFilterColumn | null) => Promise<void>
}

interface StockValueParams {
  subUnitId: string
  iterateRange: FilterValueTreeInput['iterateRange']
  filters: boolean
  filterColumn?: { filterBy?: number; filteredOutRows?: ReadonlySet<number> | null }
  alreadyChecked: string[]
  blankChecked: boolean
  filteredOutRowsByOtherColumns: number[]
}

interface Installed {
  target?: Target
}
const installations = new WeakMap<UniverRuntime, Installed>()
const patchedModels = new WeakSet<FilterModel>()

/** Reuse Univer's anchored popup, permission checks, sort actions and
 * condition editor. Only the streamed value checklist needs a new view. */
export function prepareLargeFilterPanel(
  runtime: UniverRuntime,
  ref: { current: LazyWorkbookState | null },
  target: Target,
): void {
  const installed = installations.get(runtime)
  if (installed) {
    installed.target = target
    return
  }
  const injector = runtime.univer.__getInjector()
  const manager = injector.get(ComponentManager)
  const Stock = manager.get(PANEL_KEY) as ComponentType<Record<string, unknown>> | undefined
  if (!Stock) return

  // Univer keeps this identifier private; redi returns the registered token
  // when called with the same name (and logs that it already exists).
  const log = console.error
  let identifier
  try {
    console.error = () => undefined
    identifier = createIdentifier(VALUES_SERVICE)
  } finally {
    console.error = log
  }
  if (injector.has(identifier)) return
  const locale = injector.get(LocaleService)
  const commands = injector.get(ICommandService)
  const setup: Installed = { target }

  injector.add([
    identifier,
    {
      useValue: {
        async getFilterValues(params: StockValueParams) {
          const workbook = runtime.univerAPI.getActiveWorkbook()
          const sheet = workbook?.getSheetBySheetId(params.subUnitId)
          const state = ref.current
          // The stock controller constructs its value model even when our
          // checklist is displayed. Do not make it walk a million grid rows.
          if (!workbook || !sheet || (state && isLargeFilterSheet(state, params.subUnitId))) {
            return { filterTreeItems: [], filterTreeMapCache: new Map() }
          }
          return buildFilterValueTree({
            source: sheet.getSheet() as unknown as FilterValueSource,
            iterateRange: params.iterateRange,
            styles: workbook.getWorkbook().getStyles(),
            filteredOutRowsByOtherColumns: new Set(params.filteredOutRowsByOtherColumns),
            hasFilters: params.filters,
            filterBy: params.filterColumn?.filterBy,
            filteredOutCount: params.filterColumn?.filteredOutRows?.size ?? 0,
            alreadyChecked: new Set(params.alreadyChecked),
            blankChecked: params.blankChecked,
            t: (key) => locale.t(key),
          })
        },
      },
    },
  ])
  manager.delete(PANEL_KEY)
  manager.register(PANEL_KEY, function StreamedFilterPopup(props: Record<string, unknown>) {
    const active = setup.target
    if (
      !active ||
      ref.current !== active.state ||
      runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() !==
        active.model.subUnitId
    ) {
      return <Stock {...props} />
    }
    return (
      <FilterModes
        key={`${active.model.subUnitId}:${active.column}`}
        target={active}
        commands={commands}
        locale={locale}
        stock={<Stock {...props} />}
      />
    )
  })
  commands.onCommandExecuted(({ id }) => {
    if (id === CLOSE_PANEL) delete setup.target
  })
  installations.set(runtime, setup)
}

export function streamedFilterTarget(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  sheet: UniverWorksheet,
  column: number,
  apply: Target['apply'],
): Target | undefined {
  const model = filterModelFor(runtime, sheet)
  if (!model) return
  if (!patchedModels.has(model)) {
    // Native aggregation applies the other columns' criteria itself. The
    // placeholder model does not need a JS copy of every excluded row.
    model.getFilteredOutRowsExceptCol = () => new Set()
    patchedModels.add(model)
  }
  return { state, model, column, apply }
}

function FilterModes({
  target,
  stock,
  commands,
  locale,
}: {
  target: Target
  stock: React.ReactNode
  commands: ICommandService
  locale: LocaleService
}) {
  const filter = target.model.getFilterColumn(target.column)?.getColumnData()
  const [mode, setMode] = useState(
    filter?.customFilters
      ? FilterBy.CONDITIONS
      : filter?.colorFilters
        ? FilterBy.COLORS
        : FilterBy.VALUES,
  )
  useEffect(() => {
    const listener = commands.onCommandExecuted(({ id, params }) => {
      if (id === CHANGE_BY) setMode((params as { filterBy: FilterBy }).filterBy)
    })
    return () => listener.dispose()
  }, [commands])
  if (mode !== FilterBy.VALUES) return stock
  return <ValuesPanel target={target} commands={commands} locale={locale} />
}

interface Selection {
  all: boolean
  values: ReadonlySet<string>
}
const selected = (selection: Selection, text: string) =>
  selection.all !== selection.values.has(text)

function ValuesPanel({
  target,
  commands,
  locale,
}: {
  target: Target
  commands: ICommandService
  locale: LocaleService
}) {
  const { state, model, column } = target
  const t = (key: string) => locale.t(`sheets-filter-ui.panel.${key}`)
  const toggleLabel = t('select-all')
  const criteria = model.getFilterColumn(column)?.getColumnData()
  const [selection, setSelection] = useState<Selection>(() => ({
    all: !criteria,
    values: new Set([
      ...(criteria?.filters?.filters ?? []),
      ...(criteria?.filters?.blank ? [''] : []),
    ]),
  }))
  const [query, setQuery] = useState('')
  const [offset, setOffset] = useState(0)
  const [result, setResult] = useState<{ query: string; offset: number; page: FilterPage }>()
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const mounted = useRef(true)
  const allBox = useRef<HTMLInputElement>(null)
  const list = useRef<HTMLDivElement>(null)
  const components = useComponentsOfPart(SheetsUIPart.FILTER_PANEL_EMBED_POINT)
  const close = () => void commands.executeCommand(CLOSE_PANEL)
  const page = result?.query === query && result.offset === offset ? result.page : undefined

  useEffect(() => {
    mounted.current = true
    return () => {
      mounted.current = false
    }
  }, [])
  useEffect(() => {
    let active = true
    const request = async () => {
      try {
        const page = await readLargeFilterPage(
          state,
          model,
          column,
          query,
          offset,
          PAGE_SIZE,
          () => active,
        )
        if (active && page) setResult({ query, offset, page })
      } catch (error: unknown) {
        if (active) setError(String(error))
      }
    }
    const timer = setTimeout(() => void request(), query ? 150 : 0)
    return () => {
      active = false
      clearTimeout(timer)
    }
  }, [state, model, column, query, offset])
  useEffect(() => {
    if (allBox.current) allBox.current.indeterminate = selection.values.size > 0
  }, [selection])
  useEffect(() => {
    if (list.current) list.current.scrollTop = 0
  }, [page])

  async function apply(clear = false) {
    setBusy(true)
    setError('')
    try {
      if (clear || (selection.all && selection.values.size === 0)) {
        await target.apply(null)
        return
      }
      let picked: ReadonlySet<string> = selection.values
      if (selection.all) {
        const included = new Set<string>()
        let offset = 0
        while (mounted.current) {
          const page = await readLargeFilterPage(
            state,
            model,
            column,
            '',
            offset,
            PAGE_SIZE,
            () => mounted.current,
          )
          if (!page) return
          if (page.total > MAX_BULK_VALUES) {
            setError(largeFilterSelectValuesMessage())
            return
          }
          for (const value of page.values)
            if (selected(selection, value.text)) included.add(value.text)
          offset += page.values.length
          if (offset >= page.total || !page.values.length) break
        }
        picked = included
      }
      if (!mounted.current || !picked.size) return
      await target.apply({
        colId: column,
        filters: {
          filters: [...picked].filter((value) => value !== ''),
          ...(picked.has('') ? { blank: true } : {}),
        },
      })
    } catch (error: unknown) {
      if (mounted.current) setError(String(error))
    } finally {
      if (mounted.current) setBusy(false)
    }
  }

  function title(value: FilterPage['values'][number]) {
    if (!value.text) return t('empty')
    const pattern =
      value.styleIndex === undefined ? undefined : state.file.styles[value.styleIndex]?.numberFormat
    if (pattern && numfmt.getFormatInfo(pattern)?.isDate && Number.isFinite(Number(value.text))) {
      return numfmt.format(pattern, Number(value.text))
    }
    return value.text
  }

  return (
    <div
      data-u-comp="sheets-filter-panel"
      className="streamed-filter-panel"
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.stopPropagation()
          close()
        }
      }}
    >
      <ComponentContainer
        components={components}
        sharedProps={{ range: model.getRange(), colIndex: column, onClose: close }}
      />
      <div className="streamed-filter-tabs">
        {(
          [
            ['by-values', FilterBy.VALUES],
            ['by-colors', FilterBy.COLORS],
            ['by-conditions', FilterBy.CONDITIONS],
          ] as const
        ).map(([key, value]) => (
          <button
            key={key}
            type="button"
            aria-pressed={value === FilterBy.VALUES}
            onClick={() => void commands.executeCommand(CHANGE_BY, { filterBy: value })}
          >
            {t(key)}
          </button>
        ))}
      </div>
      <input
        className="streamed-filter-search"
        type="search"
        aria-label={t('search-placeholder')}
        placeholder={t('search-placeholder')}
        value={query}
        disabled={busy}
        onChange={(event) => {
          setQuery(event.currentTarget.value)
          setOffset(0)
          setError('')
        }}
      />
      <div className="streamed-filter-checklist" data-u-comp="sheets-filter-panel-content">
        <label>
          <input
            type="checkbox"
            ref={allBox}
            checked={selection.all && !selection.values.size}
            disabled={busy || !result}
            onChange={(event) =>
              setSelection({ all: event.currentTarget.checked, values: new Set() })
            }
          />
          {toggleLabel}
        </label>
        <div className="streamed-filter-items" ref={list} aria-busy={!page && !error}>
          {(page?.values ?? []).map((value) => (
            <label key={value.text}>
              <input
                type="checkbox"
                checked={selected(selection, value.text)}
                disabled={busy}
                onChange={() =>
                  setSelection((previous) => {
                    const values = new Set(previous.values)
                    if (values.has(value.text)) values.delete(value.text)
                    else values.add(value.text)
                    return { all: previous.all, values }
                  })
                }
              />
              <span title={title(value)}>{title(value)}</span>
              <small>({value.count})</small>
            </label>
          ))}
        </div>
        {(result?.page.total ?? 0) > PAGE_SIZE && (
          <div className="streamed-filter-pages">
            <button
              type="button"
              aria-label={appText('appPrevious')}
              disabled={!page || offset === 0 || busy}
              onClick={() => setOffset(offset - PAGE_SIZE)}
            >
              ‹
            </button>
            <span>
              {page
                ? `${offset + 1}–${offset + page.values.length} / ${page.total.toLocaleString()}`
                : '…'}
            </span>
            <button
              type="button"
              aria-label={appText('appNext')}
              disabled={!page || offset + PAGE_SIZE >= page.total || busy}
              onClick={() => setOffset(offset + PAGE_SIZE)}
            >
              ›
            </button>
          </div>
        )}
      </div>
      {error && <p role="alert">{error}</p>}
      <div className="streamed-filter-footer" data-u-comp="sheets-filter-panel-footer">
        <button type="button" disabled={!criteria || busy} onClick={() => void apply(true)}>
          {t('clear-filter')}
        </button>
        <div>
          <button type="button" onClick={close}>
            {t('cancel')}
          </button>
          <button
            type="button"
            className="streamed-filter-confirm"
            disabled={!result || busy || (!selection.all && !selection.values.size)}
            onClick={() => void apply()}
          >
            {t('confirm')}
          </button>
        </div>
      </div>
    </div>
  )
}
