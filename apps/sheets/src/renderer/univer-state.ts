/**
 * Shared Univer runtime types and lazy-workbook streaming state.
 *
 * Used by both the App component (App.tsx) and the module-level sync
 * helpers (univer-sync.ts).
 */
import { BorderType, type IRange } from '@univerjs/core'

import type { WorkbookFile, WorkbookPivotDefinition } from '../shared/desktop-api'
import type { createUniver } from './create-univer'
import type { EditJournal } from './edit-journal'

export type UniverRuntime = ReturnType<typeof createUniver>
export type ActiveWorkbook = NonNullable<
  ReturnType<UniverRuntime['univerAPI']['getActiveWorkbook']>
>
export type UniverWorksheet = NonNullable<ReturnType<ActiveWorkbook['getActiveSheet']>>

export interface LazyWorkbookState {
  readonly file: WorkbookFile
  readonly generation: number
  readonly loadedRanges: Map<string, IRange>
  readonly loadingKeys: Map<string, string>
  readonly retryTimers: Map<string, ReturnType<typeof setTimeout>>
  readonly appliedMerges: Map<string, Set<string>>
  readonly appliedRowKeys: Map<string, Set<string>>
  readonly appliedCfSheets: Set<string>
  readonly appliedFilterSheets: Set<string>
  readonly appliedDvSheets: Set<string>
  /// File-side worksheet protection, known once a sheet finishes indexing.
  readonly sheetProtections: Map<string, { protected: boolean; hasPassword: boolean }>
  /// Defined names the Univer engine rejected at install — preserved verbatim
  /// by the declarative defined-names save.
  readonly uninstalledDefinedNames: Set<string>
  readonly hyperlinkTargets: Map<string, Map<string, string>>
  readonly frozenStripKeys: Map<string, string>
  /// Where each sheet's filter button came from: the worksheet's own
  /// autoFilter (saveable) or a table part (whose filter lives in the table
  /// XML — editing it is blocked).
  readonly filterOrigins: Map<string, { origin: 'worksheet' | 'table'; range: IRange }>
  /// Sheets whose view shows formulas instead of values
  /// (sheetView/@showFormulas): seeded from the file, flipped by the
  /// Formulas-tab toggle, applied
  /// to the global raw-formula render key on sheet activation.
  readonly showFormulaSheets: Set<string>
  /// Small workbooks are fully loaded with formulas handed to Univer's engine
  /// for live recalculation; large ones stream cached values only.
  readonly formulaMode: boolean
  readonly editJournal: EditJournal
  readonly flags: { preloadComplete: boolean }
  /// Closure mode: on streamed workbooks whose formula dependency closure is
  /// small, the closure cells are installed once and pinned (re-applied after
  /// viewport eviction) so the engine recalculates them live.
  readonly closure: {
    status: 'idle' | 'pending' | 'active' | 'unavailable'
    readonly pinned: Map<string, Map<string, PinnedClosureCell>>
  }
  /// Formula text per sheet ('row:col', file coordinates) harvested from
  /// readWorkbookFormulas, so the formula bar can show formulas even when the
  /// closure gave up and the engine never sees them. Display-only.
  readonly formulaText: Map<string, Map<string, string>>
  /// Parsed pivot definitions keyed by part path, loaded eagerly at open so
  /// pivot refresh stays synchronous.
  readonly pivotDefinitions: Map<string, WorkbookPivotDefinition>
  /// Known row/column outline levels per sheet (file reads + this session's
  /// group edits). Rows outside loaded ranges default to level 0.
  readonly outline: Map<
    string,
    {
      readonly rows: Map<number, { level: number; collapsed: boolean }>
      readonly cols: Map<number, { level: number; collapsed: boolean }>
    }
  >
  /// IronCalc fallback when closure mode is unavailable: file formula-cell
  /// keys per sheet, engine values overlaid on viewport patches, and a
  /// session kill switch after the engine rejects the workbook repeatedly.
  readonly recalc: {
    timer: ReturnType<typeof setTimeout> | null
    generation: number
    /// consecutive engine failures; a success resets it
    failures: number
    readonly formulaCells: Map<string, ReadonlySet<number>>
    readonly overlay: Map<string, Map<string, PinnedClosureCell>>
  }
}

export interface PinnedClosureCell {
  readonly f?: string
  readonly v?: string | number | boolean | null
}

/// Budget for closure mode: formula cells plus every precedent they read.
export const CLOSURE_MAX_CELLS = 50_000

/// Streaming re-installs viewport cells through the same mutation user edits
/// produce; this flag keeps programmatic patches out of the edit journal.
/// Shared mutable state between App.tsx and univer-sync.ts.
export const journalSuppression = { active: false }

export const BORDER_COMMAND_TYPES: Record<string, BorderType> = {
  all: BorderType.ALL,
  outer: BorderType.OUTSIDE,
  'thick-outer': BorderType.OUTSIDE,
  top: BorderType.TOP,
  bottom: BorderType.BOTTOM,
  left: BorderType.LEFT,
  right: BorderType.RIGHT,
  none: BorderType.NONE,
}
