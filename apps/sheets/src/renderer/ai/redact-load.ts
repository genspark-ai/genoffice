/**
 * Turning what the main process read out of `xl/gxRedactions.json` into the
 * index the read tools ask about a cell (see `./redact`).
 *
 * ## The one decision here
 *
 * A part that exists but cannot be read must not become an empty index. An
 * empty index answers "no cell is withheld" to every read the model makes, so a
 * damaged part read as "nothing hidden" hands over exactly the values the
 * reader chose to hide — a failure that is silent, total, and in the worst
 * direction. So:
 *
 * - `absent` (the common case) and `ok` both resolve to an index built from the
 *   part, and the renderer's message bar stays quiet.
 * - `unreadable` resolves to [`WITHHELD_UNKNOWN`], an index that withholds
 *   every cell with a neutral label, plus the reason for the user. The model
 *   then sees placeholders instead of data, which is a recoverable complaint;
 *   seeing the data would not be.
 *
 * The reason is kept out of the label on purpose: a label reaches the model, so
 * it must not report the workbook's internals back to it.
 */
import {
  buildRedactionIndex,
  NO_REDACTIONS,
  type RedactionIndex,
  type SheetRefLike,
} from './redact'
import type { WorkbookRedactionsResult } from '../../shared/desktop-api'
import type { TFunc } from '../i18n/locale'

/** The label shown for a cell whose mark could not be read. */
export const UNKNOWN_REDACTION_LABEL = 'private'

/**
 * Withhold everything, because nothing is known.
 *
 * `isEmpty` is false on purpose: it is what the readers and the write guard ask
 * first, and a truthful "this workbook withholds something" is what keeps them
 * from treating the session as unprotected.
 */
export const WITHHELD_UNKNOWN: RedactionIndex = {
  isEmpty: false,
  sheetIds: [],
  labelAt: () => UNKNOWN_REDACTION_LABEL,
  labelsFor: () => [],
  marksFor: () => [],
}

/** What the open path installs for a session. */
export interface RedactionSession {
  readonly index: RedactionIndex
  /** the reason to show the user, or null when the part was read or absent */
  readonly error: string | null
  /**
   * Something the user should know but that does not withhold anything, or
   * null when nothing happened. The one case today: the marks came from the
   * defined-name mirror because the package part is gone, which means the file
   * has been through a program that rebuilds the package. Saving here puts the
   * part back, so the reader is told rather than left guessing why.
   */
  readonly notice: string | null
}

/**
 * Resolve one open's redaction state into the index the readers share.
 *
 * The whole batch is checked before it is used, so a refusal never leaves a
 * half-applied index behind: a session gets either its real marks or nothing at
 * all to reveal.
 *
 * `t` is the app's translate function: the notice is product text, so it cannot
 * be built here from an English literal.
 */
export function redactionSessionFor(
  result: WorkbookRedactionsResult,
  sheets: readonly SheetRefLike[],
  t: TFunc,
): RedactionSession {
  if (result.status === 'unreadable') {
    return { index: WITHHELD_UNKNOWN, error: result.error, notice: null }
  }
  if (result.status === 'absent') return { index: NO_REDACTIONS, error: null, notice: null }
  return {
    index: buildRedactionIndex(result.states, sheets),
    error: null,
    notice: result.source === 'mirror' ? t('redactRecoveredFromNames') : null,
  }
}
