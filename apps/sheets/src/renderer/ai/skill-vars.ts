/**
 * Context variables for the skills attached to one message: the sheets slice
 * of what `skillContextVars` standardizes (APP, LANGUAGE, TODAY, FILE_NAME,
 * SELECTION), built at send time in front of `withAttachedSkills`.
 *
 * SELECTION is the run's frozen scope — the same snapshot the composer's scope
 * chip shows and get_workbook_context later reports — not the live grid, so a
 * skill body and the run it rides never disagree about what "this column"
 * means. A scope the user dropped, or never made (a resting single-cell
 * selection carries no intent), omits SELECTION exactly like the chip.
 *
 * The scope's values ride along only while the range is small, and go through
 * the redaction-aware readCells the model's own tools use, so a withheld cell
 * shows its placeholder instead of its secret.
 */
import { skillContextVars, type SkillVars } from '@genoffice/agent-core'
import {
  formatAddress,
  parseRange,
  rangeAddresses,
  rangeCellCount,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { FrozenSelection } from './tools'

/** Past this cell count the values stay unread: the address alone names the
 *  target, and the cells are one read_range away for the model. */
const MAX_VALUE_CELLS = 64
/** SELECTION's whole budget, address included — a context var is a hint, not a
 *  data dump. */
const MAX_SELECTION_CHARS = 500

/** One cell as readCells reports it; a withheld cell arrives with its
 *  placeholder as the value. */
export interface SelectionCell {
  readonly value: CellScalar
  readonly formula?: string | undefined
}

/** The redaction-aware cell reader the panel already hands the model's tools. */
export type ReadSelectionValues = (
  addresses: readonly string[],
  sheetId: string,
) => Record<string, SelectionCell>

export interface SheetSkillVarsInput {
  readonly lang: string
  readonly app: string
  readonly fileName?: string | undefined
  /** the run's frozen scope: null = the user dropped the chip, undefined = no
   *  run owns one — both omit SELECTION */
  readonly scope: FrozenSelection | null | undefined
  /** display name for the scope's sheet; an unknown id falls back to the raw id */
  readonly sheetNameOf: (sheetId: string) => string | undefined
  readonly readValues: ReadSelectionValues
  /** clock override for tests */
  readonly now?: Date | undefined
}

/** the context variables every skill attached to a sheets message expands against */
export function sheetSkillVars(input: SheetSkillVarsInput): SkillVars {
  const selection = selectionText(input)
  return skillContextVars({
    lang: input.lang,
    app: input.app,
    ...(input.fileName ? { fileName: input.fileName } : {}),
    ...(selection ? { selection } : {}),
    ...(input.now ? { now: input.now } : {}),
  })
}

/** `Sheet1!B2:D5` plus the values when the range is small, capped whole. */
function selectionText(input: SheetSkillVarsInput): string | undefined {
  const scope = input.scope
  if (!scope) return undefined
  const sheet = input.sheetNameOf(scope.sheetId) ?? scope.sheetId
  const text = `${sheet}!${scope.a1}${valuesSuffix(scope, input.readValues)}`
  return text.length > MAX_SELECTION_CHARS ? `${text.slice(0, MAX_SELECTION_CHARS - 1)}…` : text
}

/** `: B2=Q1, C2=1200; B3=Q2 …` — row-major, empty cells skipped, a formula in
 *  parentheses behind its computed value the way read_range reports one. */
function valuesSuffix(scope: FrozenSelection, readValues: ReadSelectionValues): string {
  let bounds: RangeBounds
  try {
    bounds = parseRange(scope.a1)
  } catch {
    return ''
  }
  if (rangeCellCount(bounds) > MAX_VALUE_CELLS) return ''
  let cells: Record<string, SelectionCell>
  try {
    // the workbook can be tearing down while the send races it; the address
    // alone still names the target
    cells = readValues(rangeAddresses(bounds), scope.sheetId)
  } catch {
    return ''
  }
  const rows: string[] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    const rowCells: string[] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const address = formatAddress(row, column)
      const text = cellText(cells[address])
      if (text !== null) rowCells.push(`${address}=${text}`)
    }
    if (rowCells.length > 0) rows.push(rowCells.join(', '))
  }
  return rows.length > 0 ? `: ${rows.join('; ')}` : ''
}

function cellText(cell: SelectionCell | undefined): string | null {
  if (cell === undefined) return null
  if (cell.formula) {
    const value = cell.value == null ? '' : String(cell.value)
    return value ? `${value} (${cell.formula})` : cell.formula
  }
  if (cell.value == null || cell.value === '') return null
  return String(cell.value)
}
