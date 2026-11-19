/**
 * The visible half of a withheld cell: a fill, so the grid says which cells the
 * model cannot see.
 *
 * ## Why this touches the reader's formatting at all
 *
 * The other four apps can show a mark without changing anything: an underlined
 * run, a `data-gx-redact` attribute, a line in the source. A spreadsheet cell
 * has no room for a label without replacing the value, and the value is the
 * reader's own data. So the mark tints the cell instead — which is a real
 * formatting change that goes into their file, and is the one thing about this
 * feature that is not purely additive.
 *
 * ## What makes that acceptable
 *
 * The colour a cell already had is carried in the mark itself
 * (`RedactionMark.previousFill`), so clearing puts it back instead of blanking
 * it. A mark therefore never silently discards the reader's own palette, and
 * undoing a mark by hand leaves the file as it was.
 *
 * ## What it does not buy
 *
 * The fill is a *visual* cue, not part of the guarantee. The model reads the
 * value through the projection in ./ai/redact, not through the fill, so a
 * reader who clears the colour by hand still gets the protection. The reverse
 * is also true and worth stating: an undone fill leaves the grid unmarked while
 * the mark — and the protection — is still there. Re-toggling the mark restores
 * the two.
 */
import { NO_FILL_STYLE } from './edit-journal'
import type { UniverRuntime, UniverWorksheet } from './univer-state'
import type { MarkKey } from './redact-actions'
import { normalizeHexColor } from './selection-format'

/**
 * The tint a withheld cell carries.
 *
 * The same purple family the other apps mark redactions with (markdown's
 * `--redact-ink` is #6b4fc0), so one gesture reads the same in every app.
 * Light enough that the cell's own black text stays readable — the value is
 * still the reader's, and they have to be able to check it.
 */
export const MARK_FILL = '#E8E1F5'

/** The range facade the grid hands out, derived rather than imported so the
 *  tint and the grid can never disagree about which Univer this is. */
type FacadeRange = ReturnType<UniverWorksheet['getRange']>

/** The fill a cell has right now, or null when it has none. */
export function readCellFill(range: FacadeRange): string | null {
  try {
    return normalizeHexColor(range.getCellStyleData()?.bg?.rgb ?? null)
  } catch {
    // A cell the grid has never touched can throw on style resolution; that is
    // "no fill", not a reason to refuse the mark.
    return null
  }
}

/**
 * Paint the mark's tint over a range, returning the fill it displaced.
 *
 * Read-then-paint lives in one function on purpose. The two orders are
 * indistinguishable in the result — both leave the tinted cell looking right —
 * and getting it backwards stores the tint as the reader's own colour, so
 * clearing paints the tint back over a cell they had coloured. Making the order
 * a property of this function rather than of its caller is what stops that.
 */
export function tintRange(range: FacadeRange, fill: string = MARK_FILL): string | null {
  const previousFill = readCellFill(range)
  range.setBackground(fill)
  return previousFill
}

/**
 * Build the mark for a range, tinting the cells as it goes.
 *
 * The previous fill is produced here and nowhere else. A caller that assembled
 * a mark itself and forgot the field would paint the tint and lose the reader's
 * colour for good, and the result looks correct right up until they clear the
 * mark — so the field is not a thing a call site can leave out.
 */
export function buildMark(
  worksheet: UniverWorksheet,
  area: string,
  label: string,
  bounds: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): MarkKey & { readonly label: string; readonly previousFill: string | null } {
  const range = worksheet.getRange(area)
  const previousFill = tintRange(range)
  return { ...bounds, label, previousFill }
}

/**
 * Put back the fill a mark displaced.
 *
 * Goes through set-style with the empty-rgb sentinel rather than
 * `setBackground(null)`: the null form reaches the mutation as a MISSING bg,
 * and a `<col style=>` fill on the column composes straight back through the
 * cell, so the "cleared" cell would look unchanged.
 */
export function restoreCellFill(
  runtime: UniverRuntime,
  sheetId: string,
  range: FacadeRange,
  previousFill: string | null | undefined,
  workbookId: string,
): void {
  if (previousFill) {
    range.setBackground(previousFill)
    return
  }
  runtime.univerAPI.syncExecuteCommand('sheet.command.set-style', {
    unitId: workbookId,
    subUnitId: sheetId,
    range: range.getRange(),
    style: { type: 'bg', value: { ...NO_FILL_STYLE } },
  })
}
