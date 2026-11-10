import type {
  BorderPatch,
  CellFormatPatch,
  ClearRangeOperation,
  ConvertToValuesOperation,
  CopyRangeOperation,
  FillRangeOperation,
  FindReplaceOperation,
  LayoutOperation,
  SortRangeOperation,
  StructuralOperation,
} from './workbook-dsl'
import type { SheetVisual } from './chart-visual'

export type CellScalar = string | number | boolean | null

export interface CellState {
  readonly value: CellScalar
  readonly formula?: string | undefined
  /** raw model value where `value` is the rendered display text (numfmt
   * interceptors rewrite dates/formatted numbers into strings); ops that must
   * not corrupt cell types read this when present */
  readonly rawValue?: CellScalar | undefined
}

/** resolved per-cell formatting; unlike CellFormatPatch, never holds nulls */
export interface CellFormatState {
  readonly bold?: boolean | undefined
  readonly italic?: boolean | undefined
  readonly underline?: boolean | undefined
  readonly strikethrough?: boolean | undefined
  readonly fontFamily?: string | undefined
  readonly fontSize?: number | undefined
  readonly fontColor?: string | undefined
  readonly fillColor?: string | undefined
  readonly numberFormat?: string | undefined
  readonly horizontalAlign?: 'left' | 'center' | 'right' | undefined
  readonly verticalAlign?: 'top' | 'center' | 'bottom' | undefined
  readonly wrapText?: boolean | undefined
  readonly textRotation?: number | 'vertical' | undefined
  readonly indent?: number | undefined
  readonly border?: BorderPatch | undefined
}

export interface WorksheetState {
  readonly id: string
  readonly name: string
  readonly cells: Readonly<Record<string, CellState>>
  /** demo-mode grid size requested via add_sheet rows/columns; the Univer
   * rebuild still grows the grid past this to fit written content */
  readonly gridRows?: number | undefined
  readonly gridColumns?: number | undefined
  /** demo-mode formatting, keyed by address; absent means "no explicit format" */
  readonly styles?: Readonly<Record<string, CellFormatState>> | undefined
  /** demo-mode merged ranges ("B2:D2"), replayed on rebuild */
  readonly merges?: readonly string[] | undefined
  /** demo-mode row heights in points, keyed by 1-based row number */
  readonly rowHeights?: Readonly<Record<string, number>> | undefined
  /** demo-mode column widths in px, keyed by column label */
  readonly colWidths?: Readonly<Record<string, number>> | undefined
  /** demo-mode charts added by AI add_chart, replayed on rebuild */
  readonly visuals?: readonly SheetVisual[] | undefined
}

export interface WorkbookSnapshot {
  readonly revision: number
  readonly sheets: readonly WorksheetState[]
}

export interface CellChange {
  readonly sheetId: string
  readonly address: string
  readonly before: CellState
  readonly after: CellState
}

export interface SheetRename {
  readonly sheetId: string
  readonly before: string
  readonly after: string
}

export interface StructuralChange {
  /** range-level bulk ops (fill_range, copy_range, convert_to_values, large
   * clear_range / find_replace / sort_range) ride here too: applied by the
   * executors, like layout ops (no per-cell before-state) */
  readonly op:
    | StructuralOperation
    | LayoutOperation
    | FillRangeOperation
    | CopyRangeOperation
    | ConvertToValuesOperation
    | ClearRangeOperation
    | FindReplaceOperation
    | SortRangeOperation
  readonly label: string
}

export interface FormatChange {
  readonly sheetId: string
  readonly range: string
  readonly format: CellFormatPatch
  readonly label: string
}

export interface ChangePlan {
  readonly transactionId: string
  readonly baseRevision: number
  readonly cellChanges: readonly CellChange[]
  readonly sheetRenames: readonly SheetRename[]
  readonly structuralChanges: readonly StructuralChange[]
  readonly formatChanges: readonly FormatChange[]
  readonly warnings: readonly string[]
}

/** result of applying a proposed plan to the live workbook */
export interface ApplyOutcome {
  readonly ok: boolean
  readonly reason?: string
  /** the failure hit mid-batch: earlier operations were already committed */
  readonly partiallyApplied?: boolean
  /** the batch exceeded the undo budget: no undo entry reverts it */
  readonly undoDropped?: boolean
  /** apply-time advisories (e.g. formulas frozen to values) surfaced in the tool result */
  readonly notices?: readonly string[]
}

export interface CommitReceipt {
  readonly transactionId: string
  readonly previousRevision: number
  readonly revision: number
}

export interface WorkbookAdapter {
  getSnapshot(): WorkbookSnapshot
  plan(input: unknown): ChangePlan
  apply(plan: ChangePlan): CommitReceipt
  undo(): CommitReceipt
}
