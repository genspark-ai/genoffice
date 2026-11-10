/**
 * Excel-style Find & Replace: glue between the app's own dialog and Univer's
 * find-replace service.
 *
 * The stock Univer dialog is a 350px panel without a results list, input
 * history or an Excel-shaped layout. The app replaces the VIEW only: sessions,
 * matching, navigation and replace stay on IFindReplaceService (including the
 * lazy-workbook provider bridge); the stock dialog stays mounted for its
 * session lifecycle but is hidden by a styles.css rule keyed on its
 * `data-u-comp="find-replace-dialog"` marker. The service exposes no way to
 * leave replace mode or to enumerate matches, so `hideReplace` and
 * `getSessionMatches` reach documented-shape internals with fail-soft guards
 * (pinned Univer 0.25.1).
 */
import type { IFindMatch, IFindReplaceService, IFindReplaceState } from '@univerjs/find-replace'
import type { IRange, Nullable, Workbook } from '@univerjs/core'
import { formatAddress, type RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'
import { scalarToText } from './lazy-find'

interface FindReplaceServiceInternals {
  _state?: { changeState(update: { replaceRevealed: boolean }): void }
  _toggleRevealReplace?: (revealed: boolean) => void
  _model?: Nullable<{ _matches?: IFindMatch[] }>
}

/**
 * Leave replace mode (the service only exposes one-way `revealReplace`).
 * `replaceRevealed` is not a research trigger, so the session and its matches
 * survive the switch. The context flag flips before the state emission so a
 * synchronous subscriber never observes the panel in replace mode with the
 * flag already cleared, or the reverse.
 */
export function hideReplace(service: IFindReplaceService): void {
  const internals = service as unknown as FindReplaceServiceInternals
  internals._toggleRevealReplace?.(false)
  internals._state?.changeState({ replaceRevealed: false })
}

/**
 * True while the Find what box holds text the session has not searched yet.
 * Replace / Replace All act on the searched string, so they must wait (or
 * search first) rather than rewrite the previous query's matches.
 */
export function hasUncommittedQuery(
  state: Pick<IFindReplaceState, 'findString' | 'inputtingFindString'>,
): boolean {
  return state.inputtingFindString !== state.findString
}

/**
 * The current session's composite match list — the same array the service's
 * `matchesCount` counts. Empty when no session is live or the internal shape
 * moved.
 */
export function getSessionMatches(service: IFindReplaceService): IFindMatch[] {
  const internals = service as unknown as FindReplaceServiceInternals
  const matches = internals._model?._matches
  return Array.isArray(matches) ? matches : []
}

/** One row of the Find All results table. */
export interface FindAllRow {
  readonly sheetId: string
  readonly sheetName: string
  readonly address: string
  readonly content: string
  readonly bounds: RangeBounds
}

/** Cell-match shape the sheets provider (and the lazy bridge) produces. */
interface SheetCellMatchLike {
  unitId: string
  range?: { subUnitId?: string; range?: IRange }
  /** Lazy out-of-window hits carry their display text along. */
  matchedText?: string | null
}

const asCellMatch = (match: IFindMatch): SheetCellMatchLike =>
  match as unknown as SheetCellMatchLike

function cellContent(workbook: Workbook, sheetId: string, row: number, column: number): string {
  const worksheet = workbook.getSheetBySheetId(sheetId)
  const cell = worksheet?.getCellRaw(row, column)
  if (!cell) return ''
  const value = scalarToText(cell.v ?? null)
  return value ?? cell.f ?? ''
}

/**
 * Shapes session matches into display rows. Matches from other units or with
 * an unexpected shape are skipped; `total` still counts every kept row so the
 * table can say how many the `limit` hid.
 */
export function buildFindAllRows(
  matches: readonly IFindMatch[],
  workbook: Workbook,
  limit: number,
): { rows: FindAllRow[]; total: number } {
  const unitId = workbook.getUnitId()
  const rows: FindAllRow[] = []
  let total = 0
  for (const match of matches) {
    const cell = asCellMatch(match)
    const subUnitId = cell.range?.subUnitId
    const range = cell.range?.range
    if (cell.unitId !== unitId || !subUnitId || !range) continue
    total += 1
    if (rows.length >= limit) continue
    const sheetName = workbook.getSheetBySheetId(subUnitId)?.getName() ?? ''
    rows.push({
      sheetId: subUnitId,
      sheetName,
      address: formatAddress(range.startRow, range.startColumn),
      content:
        cell.matchedText ?? cellContent(workbook, subUnitId, range.startRow, range.startColumn),
      bounds: {
        startRow: range.startRow,
        startColumn: range.startColumn,
        endRow: range.endRow,
        endColumn: range.endColumn,
      },
    })
  }
  return { rows, total }
}

export const FIND_HISTORY_KEY = 'genoffice.sheets.findHistory'
export const REPLACE_HISTORY_KEY = 'genoffice.sheets.replaceHistory'
export const HISTORY_LIMIT = 10

type StorageLike = Pick<Storage, 'getItem' | 'setItem'>

/** Recent inputs, newest first. Malformed or unavailable storage reads as empty. */
export function loadHistory(storage: StorageLike | undefined, key: string): string[] {
  try {
    const raw = storage?.getItem(key)
    if (!raw) return []
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return []
    return parsed
      .filter((entry): entry is string => typeof entry === 'string')
      .slice(0, HISTORY_LIMIT)
  } catch {
    return []
  }
}

/** Prepends a committed input, deduplicated, capped; returns the new list. */
export function pushHistory(
  storage: StorageLike | undefined,
  key: string,
  value: string,
): string[] {
  const trimmed = value.trim()
  const current = loadHistory(storage, key)
  if (trimmed === '') return current
  const next = [trimmed, ...current.filter((entry) => entry !== trimmed)].slice(0, HISTORY_LIMIT)
  try {
    storage?.setItem(key, JSON.stringify(next))
  } catch {
    // quota/private-mode failures just lose history persistence
  }
  return next
}
