/**
 * Expression ("use a formula to determine cells to format") conditional
 * formatting registers one dependency tree per covered cell in Univer's
 * formula engine, and the engine rebuilds that whole graph on every
 * recalculation — every viewport stream-in patch. A single whole-column rule
 * covers millions of cells, freezing the renderer for 10s+ per scroll and
 * growing the heap until the tab dies (genspark-ai/genoffice#158).
 *
 * This module shrinks what gets REGISTERED with the engine without touching
 * the rule data, the painted ranges, or the save format:
 *
 * 1. Axis folding. Almost all real-world expression rules pin the column and
 *    let the row float (`=$L2="done"` whole-row tints), so every cell of a
 *    row evaluates to the same value. When every reference in the formula is
 *    column-absolute (and no COLUMN()-style call makes the result depend on
 *    the evaluated cell's column), registering one cell per row is enough —
 *    the engine's W×H virtual trees collapse to H. Row-absolute formulas
 *    fold the other axis symmetrically.
 * 2. Row windowing. Streamed workbooks evict cell data outside the loaded
 *    viewport window, so conditional formatting outside that window has
 *    nothing to paint on (and would evaluate against evicted, blank data
 *    anyway). Rules whose registration is still huge after folding register
 *    only the loaded row window (plus margin) and lazily re-register when
 *    the stream window moves.
 *
 * Both transforms only narrow the ranges handed to
 * ConditionalFormattingFormulaService.registerFormulaWithRange; result
 * lookups (getFormulaResultWithCoords / getFormulaMatrix) are remapped so
 * every painted cell reads the value of its row's (or column's) registered
 * representative. Rule model, undo, save, and the CF panel see the original
 * ranges throughout.
 */
import type { IRange } from '@univerjs/core'
import { ConditionalFormattingFormulaService } from '@univerjs/preset-sheets-conditional-formatting'

import type { UniverRuntime } from './univer-state'

export interface CfFoldability {
  /// Every cell of a row evaluates identically (references column-absolute).
  foldColumns: boolean
  /// Every cell of a column evaluates identically (references row-absolute).
  foldRows: boolean
}

/// Below this coverage the per-cell registration is cheap enough to keep the
/// pristine engine behavior.
export const FOLD_MIN_COVERED_CELLS = 2_000
/// Registrations still larger than this after folding follow the streamed
/// row window instead of covering the whole rule.
export const WINDOW_MIN_REGISTERED_CELLS = 20_000
/// Extra rows registered around the loaded window so small scrolls don't
/// re-register.
export const WINDOW_MARGIN_ROWS = 512
/// Absolute ceiling on registered cells (binds only when no stream window
/// exists, i.e. fully-loaded in-memory workbooks with pathological rules).
export const REGISTER_HARD_CAP_CELLS = 100_000

/// Functions whose value depends on the evaluated cell's position or on
/// per-cell randomness — folding would change their semantics.
const BOTH_AXIS_BLOCKERS = /(?<![A-Z0-9_.])(INDIRECT|CELL|RAND|RANDBETWEEN|RANDARRAY)\s*\(/i
const COLUMN_AXIS_BLOCKER = /(?<![A-Z0-9_.])COLUMN\s*\(/i
const ROW_AXIS_BLOCKER = /(?<![A-Z0-9_.])ROW\s*\(/i

/// A1-style single reference: $-flags decide which axes shift per cell.
/// Not preceded by an identifier character; not followed by one, by `(`
/// (function name), or by `!` (sheet-name prefix that happens to look like a
/// ref, e.g. `Q1!A2`).
const A1_REF = /(?<![A-Za-z0-9_.$])(\$?)[A-Za-z]{1,3}(\$?)[0-9]+(?![A-Za-z0-9_(!])/g
/// Whole-column span (`A:C`, `$A:$C`): shifts horizontally unless both
/// anchors are absolute.
const COLUMN_SPAN =
  /(?<![A-Za-z0-9_.$:])(\$?)[A-Za-z]{1,3}\s*:\s*(\$?)[A-Za-z]{1,3}(?![A-Za-z0-9_(!:])/g
/// Whole-row span (`1:5`, `$1:$5`): shifts vertically unless both anchors
/// are absolute.
const ROW_SPAN = /(?<![A-Za-z0-9_.$:])(\$?)[0-9]+\s*:\s*(\$?)[0-9]+(?![A-Za-z0-9_(!:])/g
/// Leftover identifier (defined name): could hide relative references, so
/// its presence blocks folding. TRUE/FALSE literals are fine.
const BARE_IDENTIFIER = /(?<![A-Za-z0-9_.$])[A-Za-z_][A-Za-z0-9_.]*/g

/** Decide which axes of `formulaText` are shift-invariant. Conservative:
 * anything unrecognized blocks folding (worst case is the old behavior). */
export function analyzeCfFormulaFold(formulaText: string): CfFoldability {
  const none: CfFoldability = { foldColumns: false, foldRows: false }
  let text = formulaText.startsWith('=') ? formulaText.slice(1) : formulaText
  // String literals ("" escapes included) and quoted sheet prefixes can
  // contain anything — remove them before scanning for references.
  text = text.replace(/"(?:[^"]|"")*"/g, '0')
  text = text.replace(/'(?:[^']|'')*'!?/g, '')
  text = text.replace(/#[A-Za-z0-9/!?.]+/g, '0')
  // Structured table references / anything bracketed: unknown semantics.
  if (text.includes('[')) return none
  if (BOTH_AXIS_BLOCKERS.test(text)) return none
  let foldColumns = !COLUMN_AXIS_BLOCKER.test(text)
  let foldRows = !ROW_AXIS_BLOCKER.test(text)
  text = text.replace(COLUMN_SPAN, (_match, left: string, right: string) => {
    if (!left || !right) foldColumns = false
    return '0'
  })
  text = text.replace(ROW_SPAN, (_match, left: string, right: string) => {
    if (!left || !right) foldRows = false
    return '0'
  })
  text = text.replace(A1_REF, (_match, colAbs: string, rowAbs: string) => {
    if (!colAbs) foldColumns = false
    if (!rowAbs) foldRows = false
    return '0'
  })
  for (const match of text.matchAll(BARE_IDENTIFIER)) {
    const token = match[0]
    if (/^(TRUE|FALSE)$/i.test(token)) continue
    const rest = text.slice(match.index + token.length)
    const next = /^\s*(.)/.exec(rest)?.[1]
    // Function call or sheet-name prefix — handled by the scans above.
    if (next === '(' || next === '!') continue
    return none
  }
  return { foldColumns, foldRows }
}

export const sortRangesTopLeft = (ranges: readonly IRange[]): IRange[] =>
  [...ranges].sort((a, b) =>
    a.startRow !== b.startRow ? a.startRow - b.startRow : a.startColumn - b.startColumn,
  )

export const countCells = (ranges: readonly IRange[]): number =>
  ranges.reduce(
    (sum, r) => sum + (r.endRow - r.startRow + 1) * (r.endColumn - r.startColumn + 1),
    0,
  )

export interface RowWindow {
  startRow: number
  endRow: number
}

/** Narrow `sortedRanges` to what actually needs engine registration.
 * The first registered range must keep the sorted-first top-left corner:
 * the engine anchors every cell's formula offset there, so losing it would
 * shift every predicate. */
export function buildRegisteredRanges(
  sortedRanges: readonly IRange[],
  fold: CfFoldability,
  window: RowWindow | undefined,
): IRange[] {
  let ranges: IRange[] = sortedRanges.map((r) => ({
    startRow: r.startRow,
    endRow: fold.foldRows ? r.startRow : r.endRow,
    startColumn: r.startColumn,
    endColumn: fold.foldColumns ? r.startColumn : r.endColumn,
  }))
  if (window && countCells(ranges) > WINDOW_MIN_REGISTERED_CELLS) {
    const lo = Math.max(0, window.startRow - WINDOW_MARGIN_ROWS)
    const hi = window.endRow + WINDOW_MARGIN_ROWS
    const clamped: IRange[] = []
    for (const r of ranges) {
      const startRow = Math.max(r.startRow, lo)
      const endRow = Math.min(r.endRow, hi)
      if (startRow <= endRow) clamped.push({ ...r, startRow, endRow })
    }
    ranges = clamped
    const anchor = sortedRanges[0]
    const first = ranges[0]
    if (
      anchor &&
      (!first || first.startRow !== anchor.startRow || first.startColumn !== anchor.startColumn)
    ) {
      ranges.unshift({
        startRow: anchor.startRow,
        endRow: anchor.startRow,
        startColumn: anchor.startColumn,
        endColumn: anchor.startColumn,
      })
    }
  }
  let budget = REGISTER_HARD_CAP_CELLS
  const capped: IRange[] = []
  for (const r of ranges) {
    const width = r.endColumn - r.startColumn + 1
    const area = (r.endRow - r.startRow + 1) * width
    if (area <= budget) {
      capped.push(r)
      budget -= area
      continue
    }
    const keepRows = Math.floor(budget / width)
    if (keepRows > 0) capped.push({ ...r, endRow: r.startRow + keepRows - 1 })
    break
  }
  return capped
}

interface TrackedFormula {
  unitId: string
  subUnitId: string
  cfId: string
  formulaText: string
  sortedRanges: IRange[]
  fold: CfFoldability
  /// Present when the registration follows the streamed row window.
  registeredWindow?: RowWindow | undefined
}

/** Map a lookup offset (relative to the sorted-first range's top-left, the
 * coordinate space the CF paint path uses) onto the offset of the cell that
 * was actually registered for it. */
export function remapFoldedOffset(
  entry: Pick<TrackedFormula, 'sortedRanges' | 'fold'>,
  relRow: number,
  relCol: number,
): { row: number; col: number } {
  const anchor = entry.sortedRanges[0]
  if (!anchor) return { row: relRow, col: relCol }
  const absRow = anchor.startRow + relRow
  const absCol = anchor.startColumn + relCol
  for (const range of entry.sortedRanges) {
    if (
      absRow >= range.startRow &&
      absRow <= range.endRow &&
      absCol >= range.startColumn &&
      absCol <= range.endColumn
    ) {
      const row = entry.fold.foldRows ? range.startRow : absRow
      const col = entry.fold.foldColumns ? range.startColumn : absCol
      return { row: row - anchor.startRow, col: col - anchor.startColumn }
    }
  }
  return { row: relRow, col: relCol }
}

/** Structural surface of ConditionalFormattingFormulaService this module
 * relies on — kept explicit so tests can drive a plain mock. */
export interface CfFormulaServiceLike {
  registerFormulaWithRange(
    unitId: string,
    subUnitId: string,
    cfId: string,
    formulaText: string,
    ranges?: IRange[],
  ): void
  getFormulaResultWithCoords(
    unitId: string,
    subUnitId: string,
    cfId: string,
    formulaText: string,
    row?: number,
    col?: number,
  ): { status: unknown; result?: unknown }
  getFormulaMatrix(
    unitId: string,
    subUnitId: string,
    cfId: string,
    formulaText: string,
  ): { status: unknown; result?: { getValue(row: number, col: number): unknown } }
  deleteCache(unitId: string, subUnitId: string, cfId: string, formulaText?: string): unknown[]
  createCFormulaId(cfId: string, formulaText: string): string
  getSubUnitFormulaMap(
    unitId: string,
    subUnitId: string,
  ): { getValue(key: string, keys: string[]): { formulaId?: string } | undefined } | undefined
}

interface ActiveInstall {
  service: CfFormulaServiceLike
  tracked: Map<string, TrackedFormula>
}

/// One workbook per renderer tab, so windows key on sheet id alone.
const streamWindows = new Map<string, RowWindow>()
let activeInstall: ActiveInstall | null = null
/// Re-registrations queued for the next macrotask (coalesced per formula) so
/// they never run inside the stream-patch mutation flow.
const pendingReregisters = new Set<string>()

const trackKey = (unitId: string, subUnitId: string, cfFormulaId: string): string =>
  `${unitId}|${subUnitId}|${cfFormulaId}`

/** Streamed viewport patches report their loaded row window here; windowed
 * registrations follow it lazily. */
export function notifyCfStreamWindow(subUnitId: string, startRow: number, endRow: number): void {
  streamWindows.set(subUnitId, { startRow, endRow })
  const install = activeInstall
  if (!install) return
  for (const [key, entry] of install.tracked) {
    if (entry.subUnitId !== subUnitId || !entry.registeredWindow) continue
    const covered =
      startRow >= entry.registeredWindow.startRow && endRow <= entry.registeredWindow.endRow
    if (covered || pendingReregisters.has(key)) continue
    pendingReregisters.add(key)
    setTimeout(() => {
      pendingReregisters.delete(key)
      const current = activeInstall
      if (!current || current.tracked.get(key) !== entry) return
      reregister(current, entry)
    }, 0)
  }
}

export function resetCfStreamWindows(): void {
  streamWindows.clear()
}

function reregister(install: ActiveInstall, entry: TrackedFormula): void {
  const { service } = install
  const mapEntry = service
    .getSubUnitFormulaMap(entry.unitId, entry.subUnitId)
    ?.getValue(service.createCFormulaId(entry.cfId, entry.formulaText), ['id'])
  const formulaId = mapEntry?.formulaId
  service.deleteCache(entry.unitId, entry.subUnitId, entry.cfId, entry.formulaText)
  if (formulaId) {
    // deleteCache only forgets the id mapping; the engine-side formula (and
    // its per-cell trees) must go too or every window move would leak one.
    const engineRegister = (
      service as unknown as {
        _registerOtherFormulaService?: {
          deleteFormula(unitId: string, subUnitId: string, formulaIds: string[]): void
        }
      }
    )._registerOtherFormulaService
    engineRegister?.deleteFormula(entry.unitId, entry.subUnitId, [formulaId])
  }
  install.tracked.delete(
    trackKey(
      entry.unitId,
      entry.subUnitId,
      service.createCFormulaId(entry.cfId, entry.formulaText),
    ),
  )
  // Re-enters the wrapped register, which re-reads the current stream window.
  service.registerFormulaWithRange(
    entry.unitId,
    entry.subUnitId,
    entry.cfId,
    entry.formulaText,
    entry.sortedRanges,
  )
}

/** Wrap the service's register + lookup methods. Exported for tests. */
export function wrapCfFormulaService(service: CfFormulaServiceLike): { dispose(): void } {
  const tracked = new Map<string, TrackedFormula>()
  const install: ActiveInstall = { service, tracked }
  activeInstall = install

  const originalRegister = service.registerFormulaWithRange.bind(service)
  const originalWithCoords = service.getFormulaResultWithCoords.bind(service)
  const originalMatrix = service.getFormulaMatrix.bind(service)
  const originalDeleteCache = service.deleteCache.bind(service)

  service.registerFormulaWithRange = (unitId, subUnitId, cfId, formulaText, ranges) => {
    if (!ranges || ranges.length === 0) {
      originalRegister(unitId, subUnitId, cfId, formulaText, ranges)
      return
    }
    const cfFormulaId = service.createCFormulaId(cfId, formulaText)
    // Already registered (the paint path re-calls this on every precompute).
    if (service.getSubUnitFormulaMap(unitId, subUnitId)?.getValue(cfFormulaId, ['id'])) return
    const sortedRanges = sortRangesTopLeft(ranges)
    if (countCells(sortedRanges) <= FOLD_MIN_COVERED_CELLS) {
      originalRegister(unitId, subUnitId, cfId, formulaText, ranges)
      return
    }
    const fold = analyzeCfFormulaFold(formulaText)
    const window = streamWindows.get(subUnitId)
    const registered = buildRegisteredRanges(sortedRanges, fold, window)
    const windowed =
      window !== undefined &&
      countCells(
        sortedRanges.map((r) => ({
          ...r,
          endRow: fold.foldRows ? r.startRow : r.endRow,
          endColumn: fold.foldColumns ? r.startColumn : r.endColumn,
        })),
      ) > WINDOW_MIN_REGISTERED_CELLS
    if (!fold.foldColumns && !fold.foldRows && !windowed) {
      // Nothing to narrow (hard cap aside) — keep pristine behavior below
      // the cap so small unfoldable rules stay untouched.
      if (countCells(sortedRanges) <= REGISTER_HARD_CAP_CELLS) {
        originalRegister(unitId, subUnitId, cfId, formulaText, ranges)
        return
      }
    }
    tracked.set(trackKey(unitId, subUnitId, cfFormulaId), {
      unitId,
      subUnitId,
      cfId,
      formulaText,
      sortedRanges,
      fold,
      registeredWindow: windowed
        ? {
            startRow: Math.max(0, (window as RowWindow).startRow - WINDOW_MARGIN_ROWS),
            endRow: (window as RowWindow).endRow + WINDOW_MARGIN_ROWS,
          }
        : undefined,
    })
    originalRegister(unitId, subUnitId, cfId, formulaText, registered)
  }

  service.getFormulaResultWithCoords = (unitId, subUnitId, cfId, formulaText, row = 0, col = 0) => {
    const entry = tracked.get(
      trackKey(unitId, subUnitId, service.createCFormulaId(cfId, formulaText)),
    )
    if (!entry) return originalWithCoords(unitId, subUnitId, cfId, formulaText, row, col)
    const mapped = remapFoldedOffset(entry, row, col)
    return originalWithCoords(unitId, subUnitId, cfId, formulaText, mapped.row, mapped.col)
  }

  service.getFormulaMatrix = (unitId, subUnitId, cfId, formulaText) => {
    const outcome = originalMatrix(unitId, subUnitId, cfId, formulaText)
    const entry = tracked.get(
      trackKey(unitId, subUnitId, service.createCFormulaId(cfId, formulaText)),
    )
    const matrix = outcome.result
    if (!entry || !matrix) return outcome
    // Wrap immutably — never patch the original instance: if the service
    // ever caches and re-returns the same matrix, re-wrapping a mutated
    // getValue would nest remap layers. Prototype delegation keeps every
    // other ObjectMatrix method intact.
    const remapped = Object.create(matrix) as typeof matrix
    remapped.getValue = (row: number, col: number) => {
      const mapped = remapFoldedOffset(entry, row, col)
      return matrix.getValue(mapped.row, mapped.col)
    }
    return { ...outcome, result: remapped }
  }

  service.deleteCache = (unitId, subUnitId, cfId, formulaText) => {
    if (formulaText === undefined) {
      // Rule removed/changed: drop every tracked formula of this cfId.
      for (const [key, entry] of tracked) {
        if (entry.unitId === unitId && entry.subUnitId === subUnitId && entry.cfId === cfId) {
          tracked.delete(key)
        }
      }
    } else {
      tracked.delete(trackKey(unitId, subUnitId, service.createCFormulaId(cfId, formulaText)))
    }
    return originalDeleteCache(unitId, subUnitId, cfId, formulaText)
  }

  return {
    dispose() {
      service.registerFormulaWithRange = originalRegister
      service.getFormulaResultWithCoords = originalWithCoords
      service.getFormulaMatrix = originalMatrix
      service.deleteCache = originalDeleteCache
      tracked.clear()
      if (activeInstall === install) activeInstall = null
    },
  }
}

/** Resolve the injector's CF formula service singleton and wrap it. */
export function installCfFormulaFold(runtime: UniverRuntime): { dispose(): void } {
  resetCfStreamWindows()
  let service: CfFormulaServiceLike
  try {
    service = runtime.univer
      .__getInjector()
      .get(ConditionalFormattingFormulaService) as unknown as CfFormulaServiceLike
  } catch {
    return { dispose() {} } // CF plugin not installed in this runtime
  }
  if (
    typeof service.registerFormulaWithRange !== 'function' ||
    typeof service.getFormulaResultWithCoords !== 'function' ||
    typeof service.getFormulaMatrix !== 'function' ||
    typeof service.createCFormulaId !== 'function'
  ) {
    return { dispose() {} }
  }
  return wrapCfFormulaService(service)
}
