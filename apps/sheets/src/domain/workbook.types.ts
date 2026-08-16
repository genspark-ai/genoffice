import type {
  BorderPatch,
  CellFormatPatch,
  LayoutOperation,
  StructuralOperation,
} from './workbook-dsl'
import type { SheetVisual } from './chart-visual'

export type CellScalar = string | number | boolean | null

export interface CellState {
  readonly value: CellScalar
  readonly formula?: string | undefined
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
  readonly op: StructuralOperation | LayoutOperation
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
