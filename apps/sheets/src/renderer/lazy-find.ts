/**
 * Full-sheet Find coverage for streamed (lazy) workbooks.
 *
 * Univer's Find dialog searches the in-memory cell matrix, but a streamed
 * workbook only holds rows that were already scrolled into view, so matches
 * in never-visited rows are invisible and users conclude the data does not
 * exist (issue #113). This module wraps the built-in sheets find provider:
 * the inner model keeps handling everything inside the loaded window, while
 * the wrapper extends the session with out-of-window matches paged from the
 * underlying file via readSheetRangeMapped (journal edits included) — the
 * same approach the AI-side find takes. Focusing an out-of-window match
 * activates its sheet, starts loading its range, scrolls to it, and selects
 * it, so the grid shows real data instead of an empty jump.
 */
import type { IRange, Workbook } from '@univerjs/core'
import {
  FindBy,
  FindModel,
  IFindReplaceService,
  type IFindMatch,
  type IFindMoveParams,
  type IFindQuery,
  type IFindReplaceProvider,
  type IReplaceAllResult,
} from '@univerjs/find-replace'
import { IUniverInstanceService } from '@univerjs/core'
import { Subject, type Subscription } from 'rxjs'
import { FILE_READ_BATCH_CELLS, MAX_SCAN_CELLS } from './ai/workbook-search'
import { t } from './i18n/locale'
import { netAxisDelta } from './view-transform'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import { ensureLazyRangeLoaded, readSheetRangeMapped } from './univer-sync'

/** Same match shape the built-in sheets provider produces (ISheetCellMatch). */
export interface LazyCellMatch extends IFindMatch {
  isFormula: boolean
  replaceable?: boolean
  /// Extra bookkeeping the wrapper needs to focus/replace the hit; ignored
  /// by Univer's composite model.
  range: { subUnitId: string; range: IRange }
  matchedText?: string | null
}

export interface LazyCellTexts {
  /** Display/computed value stringified like Univer's extractPureValue. */
  value: string | null
  formula: string | undefined
}

type LazyCellTest = (cell: LazyCellTexts) => boolean

/** Whether a find session needs the file-backed extension for this workbook. */
export function planLazyFind(state: LazyWorkbookState | null): 'inactive' | 'extend' {
  if (!state || state.flags.preloadComplete) return 'inactive'
  return 'extend'
}

/// Stringifies a scalar like Univer's extractPureValue: numbers become their
/// decimal text, booleans become "1"/"0".
export function scalarToText(value: unknown): string | null {
  if (value === null || value === undefined) return null
  if (typeof value === 'number') return `${value}`
  if (typeof value === 'boolean') return value ? '1' : '0'
  return String(value)
}

/// Mirrors the built-in provider's _preprocessQuery: lowercase unless
/// case-sensitive, then trim.
export function preprocessNeedle(query: IFindQuery): string {
  const raw = query.findString ?? ''
  return (query.caseSensitive === true ? raw : raw.toLowerCase()).trim()
}

/// Mirrors Univer's matchCellData/hitCell semantics for file-backed cells:
/// substring vs whole-cell (spaces trimmed, line breaks kept), case
/// sensitivity, and formula-vs-value look-in.
export function buildLazyCellTest(query: IFindQuery): LazyCellTest | null {
  const caseSensitive = query.caseSensitive === true
  // The built-in provider preprocesses the query once up front (lowercase
  // unless case-sensitive, then trim); mirror it so both models agree on
  // what a hit is.
  const needle = preprocessNeedle(query)
  if (!needle) return null
  const matches = (text: string | null | undefined): boolean => {
    if (text === null || text === undefined) return false
    const haystack = caseSensitive ? text : text.toLowerCase()
    if (query.matchesTheWholeCell) {
      const trimmed = haystack.replace(/^ +/g, '').replace(/ +$/g, '')
      return trimmed === needle
    }
    return haystack.includes(needle)
  }
  return ({ value, formula }) => {
    if (formula && query.findBy === FindBy.FORMULA) return matches(formula)
    return matches(value)
  }
}

function insideRange(range: IRange | undefined, row: number, column: number): boolean {
  if (!range) return false
  return (
    row >= range.startRow &&
    row <= range.endRow &&
    column >= range.startColumn &&
    column <= range.endColumn
  )
}

/** True when the loaded window covers the coordinate — the inner model owns it. */
export function coveredByWindow(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): boolean {
  return insideRange(state.loadedRanges.get(sheetId), row, column)
}

function matchKey(match: LazyCellMatch): string {
  return `${match.range.subUnitId}|${match.range.range.startRow}|${match.range.range.startColumn}`
}

/** Dedupes inner-model matches against the extension's; inner entries win. */
export function mergeFindMatches(primary: IFindMatch[], extra: LazyCellMatch[]): IFindMatch[] {
  if (extra.length === 0) return primary
  const seen = new Set(primary.map((match) => matchKey(match as LazyCellMatch)))
  const merged = [...primary]
  for (const match of extra) {
    const key = matchKey(match)
    if (seen.has(key)) continue
    seen.add(key)
    merged.push(match)
  }
  return merged
}

interface ScanCell {
  readonly row: number
  readonly column: number
  readonly value: string | number | boolean | null
  readonly formula: string | undefined
}

/** Journal edits of one sheet that sit outside the loaded window. */
export function collectJournalMatches(
  state: LazyWorkbookState,
  sheetId: string,
  test: LazyCellTest,
): ScanCell[] {
  const found: ScanCell[] = []
  const journal = state.editJournal.cells.get(sheetId)
  for (const entry of journal?.values() ?? []) {
    if (!entry.hasValue) continue
    if (coveredByWindow(state, sheetId, entry.row, entry.column)) continue
    if (!test({ value: scalarToText(entry.value), formula: entry.formula })) continue
    found.push({ row: entry.row, column: entry.column, value: entry.value, formula: entry.formula })
  }
  return found
}

/** Coordinates whose file cell is shadowed by a journal edit this session.
 *  Style-only entries (hasValue false) leave the file's content authoritative
 *  and must stay findable. */
export function journalShadowKeys(state: LazyWorkbookState, sheetId: string): Set<string> {
  const shadowed = new Set<string>()
  const journal = state.editJournal.cells.get(sheetId)
  for (const entry of journal?.values() ?? []) {
    if (entry.hasValue) shadowed.add(`${entry.row}:${entry.column}`)
  }
  return shadowed
}

/** Row-major/column-major ordering across sheets, following the query direction. */
export function extraComparator(
  sheetOrder: ReadonlyMap<string, number>,
  columnDirection: boolean,
): (a: ScanCell & { sheetId: string }, b: ScanCell & { sheetId: string }) => number {
  return (a, b) => {
    const sheetDelta = (sheetOrder.get(a.sheetId) ?? 0) - (sheetOrder.get(b.sheetId) ?? 0)
    if (sheetDelta !== 0) return sheetDelta
    return columnDirection
      ? a.column - b.column || a.row - b.row
      : a.row - b.row || a.column - b.column
  }
}

function makeCellMatch(
  unitId: string,
  sheetId: string,
  cell: ScanCell,
  findByFormula: boolean,
): LazyCellMatch {
  const isFormula = Boolean(cell.formula)
  return {
    provider: 'sheets-find-replace-provider',
    unitId,
    isFormula,
    // Formula hits are only replaced when searching formulas (mirrors the
    // built-in model); plain cells behave exactly like in-memory ones.
    replaceable: isFormula ? findByFormula : cell.value !== null && cell.value !== undefined,
    matchedText: (findByFormula && isFormula ? cell.formula : scalarToText(cell.value)) ?? null,
    range: {
      subUnitId: sheetId,
      range: {
        startRow: cell.row,
        endRow: cell.row,
        startColumn: cell.column,
        endColumn: cell.column,
      },
    },
  }
}

export interface LazyFindBridgeDeps {
  runtime: UniverRuntime
  lazyWorkbookRef: { current: LazyWorkbookState | null }
  setMessage: (message: string) => void
}

interface InnerFindModel extends FindModel {
  readonly unitId: string
  focusSelection(): void
}

/**
 * Replaces the registered sheets find provider with a wrapper while the app
 * lives. Non-streamed workbooks flow straight through; streamed ones get the
 * extended model. On dispose the adopted providers go back into the service.
 */
export function installLazyFindBridge(deps: LazyFindBridgeDeps): { dispose(): void } {
  const service = deps.runtime.univer.__getInjector().get(IFindReplaceService)
  const providers = service.getProviders()
  const adopted = new Set<IFindReplaceProvider>()
  let generation = 0
  // Registering only APPENDS to the service's live provider set, and
  // _startSearching dispatches to every provider in it — leaving the built-in
  // registered would run it twice per search (double-counted in-window
  // matches, and its second find() disposes the first session's model).
  // Detach built-ins into `adopted` so the service reaches them only through
  // the wrapper; re-sweep on every call in case one registered later.
  const adoptForeign = () => {
    for (const provider of [...providers]) {
      if (provider === wrapper) continue
      providers.delete(provider)
      adopted.add(provider)
    }
  }
  const wrapper: IFindReplaceProvider = {
    async find(query: IFindQuery) {
      generation += 1
      const liveGeneration = generation
      adoptForeign()
      const models: FindModel[] = []
      for (const builtin of [...adopted]) {
        models.push(...(await builtin.find(query)))
      }
      const state = deps.lazyWorkbookRef.current
      if (!state || planLazyFind(state) !== 'extend' || models.length === 0) return models
      return models.map(
        (model) =>
          new LazyExtendedFindModel(
            model as InnerFindModel,
            state,
            query,
            deps,
            () => generation === liveGeneration,
          ),
      )
    },
    terminate() {
      generation += 1
      adoptForeign()
      for (const builtin of adopted) builtin.terminate()
    },
  }
  const registration = service.registerFindReplaceProvider(wrapper)
  adoptForeign()
  return {
    dispose() {
      generation += 1
      registration.dispose()
      for (const builtin of adopted) providers.add(builtin)
      adopted.clear()
    },
  }
}

/**
 * A FindModel combining the built-in in-window session with out-of-window
 * hits paged from the underlying file. The inner model keeps navigating and
 * highlighting everything it can see; this wrapper only steps in when the
 * inner session runs out, and hands focus back once the jumped-to region is
 * materialized in the grid (the inner model re-runs on mutations and takes
 * over navigation there).
 */
export class LazyExtendedFindModel extends FindModel {
  readonly unitId: string

  readonly matchesUpdate$ = new Subject<IFindMatch[]>()
  readonly activelyChangingMatch$ = new Subject<LazyCellMatch>()

  private readonly state: LazyWorkbookState
  private readonly query: IFindQuery
  private readonly deps: LazyFindBridgeDeps
  private readonly isLiveGeneration: () => boolean

  private alive = true
  private truncated = false
  private extras: LazyCellMatch[] = []
  private lastFocusedExtra: LazyCellMatch | null = null
  private readonly forwardSub: Subscription

  constructor(
    private readonly inner: InnerFindModel,
    state: LazyWorkbookState,
    query: IFindQuery,
    deps: LazyFindBridgeDeps,
    isLiveGeneration: () => boolean,
  ) {
    super()
    this.state = state
    this.query = query
    this.deps = deps
    this.isLiveGeneration = isLiveGeneration
    this.unitId = inner.unitId
    // The inner model refreshes itself when grid mutations stream regions in
    // or evict them; forward those moments so the dialog count stays right.
    this.forwardSub = inner.matchesUpdate$.subscribe(() => this.emitMerged())
    void this.runScan()
  }

  override dispose(): void {
    this.alive = false
    this.forwardSub.unsubscribe()
    this.matchesUpdate$.complete()
    this.activelyChangingMatch$.complete()
    super.dispose()
  }

  getMatches(): IFindMatch[] {
    return mergeFindMatches(this.innerMatches(), this.currentExtras())
  }

  moveToNextMatch(params?: IFindMoveParams): LazyCellMatch | null {
    return this.moveThroughMatches('next', params)
  }

  moveToPreviousMatch(params?: IFindMoveParams): LazyCellMatch | null {
    return this.moveThroughMatches('previous', params)
  }

  /**
   * Segmented cursor: the inner (index-cursor) session first, then the
   * extras, then wrap back to the inner. The upstream `loop` must never
   * reach the inner model — its loop takes the modulo of its own match list
   * and would cycle in-window hits forever, starving the extras. When the
   * inner runs out it resets its own cursor, so the wrap-around re-entry
   * uses ignoreSelection to land on its first/last match. Boundary order is
   * not strict document order (a mid-sheet window hands over to the
   * top-most extra) — accepted simplification.
   */
  private moveThroughMatches(
    direction: 'next' | 'previous',
    params?: IFindMoveParams,
  ): LazyCellMatch | null {
    if (!this.lastFocusedExtra) {
      const candidate = this.innerNeighbor(direction, params)
      if (candidate) {
        if (!params?.noFocus) this.safeInnerFocus()
        return candidate as LazyCellMatch
      }
    }
    const target = this.neighborExtra(direction)
    if (target) {
      if (!params?.noFocus) this.focusExtra(target)
      else this.lastFocusedExtra = target
      return target
    }
    if (params?.loop === false) return null
    this.lastFocusedExtra = null
    const wrapped = this.innerNeighbor(direction, { ...params, ignoreSelection: true })
    if (wrapped) {
      if (!params?.noFocus) this.safeInnerFocus()
      return wrapped as LazyCellMatch
    }
    // No inner matches at all — cycle within the extras themselves.
    const candidates = this.currentExtras()
    const first =
      direction === 'next' ? (candidates[0] ?? null) : (candidates[candidates.length - 1] ?? null)
    if (!first) return null
    if (!params?.noFocus) this.focusExtra(first)
    else this.lastFocusedExtra = first
    return first
  }

  async replace(replaceString: string): Promise<boolean> {
    // Only an extra that currently holds the segmented cursor may be
    // written; otherwise the inner session owns the current match.
    const extra = this.lastFocusedExtra
    if (extra) {
      if (extra.replaceable !== true) return false
      return this.writeExtraReplacement(extra, replaceString)
    }
    try {
      return await this.inner.replace(replaceString)
    } catch {
      return false
    }
  }

  async replaceAll(replaceString: string): Promise<IReplaceAllResult> {
    let success = 0
    let failure = 0
    try {
      const result = await this.inner.replaceAll(replaceString)
      success += result.success
      failure += result.failure
    } catch {
      /* the inner session may already be gone; still report the extension's */
    }
    for (const extra of this.currentExtras()) {
      if (extra.replaceable !== true) {
        failure += 1
        continue
      }
      if (await this.writeExtraReplacement(extra, replaceString)) success += 1
      else failure += 1
    }
    return { success, failure }
  }

  focusSelection(): void {
    if (this.lastFocusedExtra) {
      this.focusExtra(this.lastFocusedExtra)
      return
    }
    this.safeInnerFocus()
  }

  private safeInnerFocus(): void {
    try {
      this.inner.focusSelection()
    } catch {
      /* closed workbook */
    }
  }

  private innerNeighbor(
    direction: 'next' | 'previous',
    params?: IFindMoveParams,
  ): IFindMatch | null {
    // loop stays stripped: the inner model's own loop cycles its list
    // forever and would never yield to the extras.
    const stripped = { ...params, noFocus: true, loop: false }
    try {
      return direction === 'next'
        ? this.inner.moveToNextMatch(stripped)
        : this.inner.moveToPreviousMatch(stripped)
    } catch {
      return null
    }
  }

  private innerMatches(): IFindMatch[] {
    try {
      return this.inner.getMatches()
    } catch {
      return []
    }
  }

  /** Extras that are still outside the (evolving) loaded window. */
  private currentExtras(): LazyCellMatch[] {
    const isRowHidden = this.rowHiddenTest()
    return this.extras.filter((match) => {
      if (
        coveredByWindow(
          this.state,
          match.range.subUnitId,
          match.range.range.startRow,
          match.range.range.startColumn,
        )
      ) {
        return false
      }
      // The in-memory scan skips rows hidden by an active filter; hold
      // out-of-window hits to the same visibility rules. Fails open when the
      // filter state is not reachable (sheet gone mid-session).
      return !isRowHidden(match.range.subUnitId, match.range.range.startRow)
    })
  }

  /// Mirrors the built-in scan's worksheet.getRowFiltered check for cells
  /// that never streamed into Univer's matrix: the filter model lives on the
  /// workbook instance, not the grid, so it answers regardless of loading.
  /// The workbook resolves once per call — currentExtras runs on every
  /// match-list read and extras can number in the thousands.
  private rowHiddenTest(): (subUnitId: string, row: number) => boolean {
    let resolved: Workbook | null = null
    try {
      resolved =
        this.deps.runtime.univer
          .__getInjector()
          .get(IUniverInstanceService)
          .getUnit<Workbook>(this.unitId) ?? null
    } catch {
      resolved = null
    }
    if (!resolved) return () => false
    return (subUnitId, row) => {
      try {
        return resolved.getSheetBySheetId(subUnitId)?.getRowFiltered(row) === true
      } catch {
        return false
      }
    }
  }

  /**
   * The first out-of-window hit after (or before) the current selection.
   * Exhaustion returns null — wrapping across the segments is
   * moveThroughMatches' job.
   */
  private neighborExtra(direction: 'next' | 'previous'): LazyCellMatch | null {
    const candidates = this.currentExtras()
    if (candidates.length === 0) return null
    const order = this.sheetOrderIndex()
    const columnDirection = this.query.findDirection === 'column'
    const axes = (row: number, column: number): [number, number] =>
      columnDirection ? [column, row] : [row, column]
    const sign = direction === 'next' ? 1 : -1
    const positionOf = (match: LazyCellMatch): [number, number, number] => {
      const bounds = match.range.range
      const [primary, secondary] = axes(bounds.startRow, bounds.startColumn)
      return [
        sign * (order.get(match.range.subUnitId) ?? Number.MAX_SAFE_INTEGER),
        sign * primary,
        sign * secondary,
      ]
    }
    const reference = this.referencePosition(order)
    if (!reference) {
      return direction === 'next' ? candidates[0]! : candidates[candidates.length - 1]!
    }
    const [referencePrimary, referenceSecondary] = axes(reference.row, reference.column)
    const referencePosition_: [number, number, number] = [
      sign * reference.sheetIndex,
      sign * referencePrimary,
      sign * referenceSecondary,
    ]
    const ordered = [...candidates].sort((a, b) => compareTriples(positionOf(a), positionOf(b)))
    return (
      ordered.find((match) => compareTriples(positionOf(match), referencePosition_) > 0) ?? null
    )
  }

  private sheetOrderIndex(): Map<string, number> {
    try {
      const sheets = this.deps.runtime.univerAPI.getActiveWorkbook()?.getSheets() ?? []
      return new Map(sheets.map((sheet, index) => [sheet.getSheetId(), index] as const))
    } catch {
      return new Map()
    }
  }

  private referencePosition(
    order: ReadonlyMap<string, number>,
  ): { sheetIndex: number; row: number; column: number } | null {
    try {
      const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
      const range = workbook?.getActiveRange()
      const activeSheet = workbook?.getActiveSheet()
      if (!workbook || !range || !activeSheet) return null
      return {
        sheetIndex: order.get(activeSheet.getSheetId()) ?? Number.MAX_SAFE_INTEGER,
        row: range.getRow(),
        column: range.getColumn(),
      }
    } catch {
      return this.lastFocusedExtra
        ? {
            sheetIndex: order.get(this.lastFocusedExtra.range.subUnitId) ?? Number.MAX_SAFE_INTEGER,
            row: this.lastFocusedExtra.range.range.startRow,
            column: this.lastFocusedExtra.range.range.startColumn,
          }
        : null
    }
  }

  /** Guards against a session outliving its workbook: after a workbook
   *  switch, sheetIds may collide and getActiveWorkbook() targets the wrong
   *  book. */
  private stateIsCurrent(): boolean {
    return this.deps.lazyWorkbookRef.current === this.state
  }

  /** Activate the sheet, load the region, scroll to it, and select the cell. */
  private focusExtra(match: LazyCellMatch): void {
    if (!this.stateIsCurrent()) return
    this.lastFocusedExtra = match
    try {
      const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
      if (!workbook) return
      const worksheet = workbook.getSheetBySheetId(match.range.subUnitId)
      if (!worksheet) return
      if (worksheet.getSheetId() !== workbook.getActiveSheet()?.getSheetId()) {
        workbook.setActiveSheet(worksheet)
      }
      const bounds = match.range.range
      // Best-effort streaming: the scroll below also triggers the regular
      // viewport load; this makes sure the exact hit lands even when the
      // visible window math picks a different anchor.
      void ensureLazyRangeLoaded(
        this.deps.runtime,
        this.deps.lazyWorkbookRef,
        worksheet,
        {
          startRow: bounds.startRow,
          endRow: bounds.endRow,
          startColumn: bounds.startColumn,
          endColumn: bounds.endColumn,
        },
        this.deps.setMessage,
      )
      worksheet.scrollToCell(bounds.startRow, bounds.startColumn)
      worksheet.getRange(bounds.startRow, bounds.startColumn, 1, 1).activate()
      this.emitMerged()
      this.activelyChangingMatch$.next(match)
    } catch {
      /* closed workbook mid-jump */
    }
  }

  /** Writes the replacement straight onto the cell; the journal carries it. */
  private async writeExtraReplacement(
    match: LazyCellMatch,
    replaceString: string,
  ): Promise<boolean> {
    if (!this.stateIsCurrent()) return false
    try {
      const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
      const worksheet = workbook?.getSheetBySheetId(match.range.subUnitId)
      if (!worksheet) return false
      const bounds = match.range.range
      const target = worksheet.getRange(bounds.startRow, bounds.startColumn, 1, 1)
      if (match.isFormula) {
        target.setValues([
          [{ f: replaceAllOccurrences(match.matchedText ?? '', this.query, replaceString) }],
        ])
      } else {
        target.setValues([
          [{ v: replaceAllOccurrences(match.matchedText ?? '', this.query, replaceString) }],
        ])
      }
      return true
    } catch {
      return false
    }
  }

  private emitMerged(): void {
    if (!this.alive) return
    this.matchesUpdate$.next(this.getMatches())
  }

  /** Pages the underlying file for out-of-window hits, emitting as it goes. */
  private async runScan(): Promise<void> {
    const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
    if (!workbook) return
    const test = buildLazyCellTest(this.query)
    if (!test) return
    const unitId = this.unitId
    const sheets = workbook.getSheets()
    const targets =
      this.query.findScope === 'unit'
        ? sheets
        : sheets.filter((sheet) => sheet.getSheetId() === workbook.getActiveSheet()?.getSheetId())
    const sheetOrder = new Map(sheets.map((sheet, index) => [sheet.getSheetId(), index] as const))
    const comparator = extraComparator(sheetOrder, this.query.findDirection === 'column')
    const collected: (ScanCell & { sheetId: string })[] = []
    // Budget counts scanned extent, not hits — the AI-side findInLazyWorkbook
    // semantics. Counting hits would scan sparse multi-million-cell sheets
    // end to end.
    let scannedCells = 0

    for (const worksheet of targets) {
      const sheetId = worksheet.getSheetId()
      // Session edits first — they shadow file cells at the same coordinates.
      for (const cell of collectJournalMatches(this.state, sheetId, test)) {
        collected.push({ ...cell, sheetId })
      }
      const meta = this.state.file.sheets.find((candidate) => candidate.id === sheetId)
      // Sheets added this session live entirely in the journal.
      if (!meta || meta.rowCount <= 0 || meta.columnCount <= 0) {
        this.refreshExtras(collected, comparator, unitId)
        this.emitMerged()
        continue
      }
      const ops = this.state.editJournal.structuralOps.get(sheetId) ?? []
      const screenRows = Math.max(meta.rowCount + netAxisDelta(ops, 'row'), 0)
      const screenColumns = Math.max(meta.columnCount + netAxisDelta(ops, 'column'), 0)
      if (screenRows <= 0 || screenColumns <= 0) continue
      const shadowed = journalShadowKeys(this.state, sheetId)
      const batchRows = Math.max(1, Math.floor(FILE_READ_BATCH_CELLS / screenColumns))
      for (let startRow = 0; startRow < screenRows; startRow += batchRows) {
        if (!this.alive || !this.isLiveGeneration() || !this.stateIsCurrent()) return
        if (this.truncated) break
        if (scannedCells >= MAX_SCAN_CELLS) {
          this.truncated = true
          break
        }
        const endRow = Math.min(startRow + batchRows - 1, screenRows - 1)
        let mapped
        try {
          mapped = await readSheetRangeMapped(
            this.state,
            sheetId,
            { startRow, endRow, startColumn: 0, endColumn: screenColumns - 1 },
            meta,
          )
        } catch {
          this.truncated = true
          break
        }
        if (!mapped) continue
        scannedCells += (endRow - startRow + 1) * screenColumns
        if (
          !mapped.raw.indexingComplete &&
          (mapped.indexedThroughScreen === null || mapped.indexedThroughScreen < endRow)
        ) {
          this.truncated = true
        }
        for (const cell of mapped.screen.cells) {
          if (shadowed.has(`${cell.row}:${cell.column}`)) continue
          if (coveredByWindow(this.state, sheetId, cell.row, cell.column)) continue
          if (!test({ value: scalarToText(cell.value), formula: cell.formula })) continue
          collected.push({
            row: cell.row,
            column: cell.column,
            value: cell.value,
            formula: cell.formula,
            sheetId,
          })
        }
        this.refreshExtras(collected, comparator, unitId)
        this.emitMerged()
      }
      if (this.truncated) break
    }

    this.refreshExtras(collected, comparator, unitId)
    if (this.truncated && this.alive && this.isLiveGeneration()) {
      // Report what was actually scanned — truncation can also come from a
      // failed read or indexing lag long before the budget.
      this.deps.setMessage(t('appFindScanTruncated', { cells: scannedCells.toLocaleString() }))
    }
    this.emitMerged()
  }

  private refreshExtras(
    collected: (ScanCell & { sheetId: string })[],
    comparator: (a: ScanCell & { sheetId: string }, b: ScanCell & { sheetId: string }) => number,
    unitId: string,
  ): void {
    const findByFormula = this.query.findBy === FindBy.FORMULA
    this.extras = [...collected]
      .sort(comparator)
      .map((cell) => makeCellMatch(unitId, cell.sheetId, cell, findByFormula))
  }
}

/// Substring replacement honoring the query's case sensitivity, replacing
/// every occurrence like Excel's Replace All.
function replaceAllOccurrences(text: string, query: IFindQuery, replaceString: string): string {
  const needle = preprocessNeedle(query)
  if (!needle) return text
  const haystack = query.caseSensitive === true ? text : text.toLowerCase()
  let result = ''
  let cursor = 0
  for (;;) {
    const index = haystack.indexOf(needle, cursor)
    if (index < 0) {
      result += text.slice(cursor)
      break
    }
    result += text.slice(cursor, index) + replaceString
    cursor = index + needle.length
  }
  return result
}

function compareTriples(a: [number, number, number], b: [number, number, number]): number {
  return a[0] - b[0] || a[1] - b[1] || a[2] - b[2]
}
