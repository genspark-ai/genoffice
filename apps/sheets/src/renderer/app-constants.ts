/**
 * App-scope constants shared by the App component: chart ribbon command
 * vocabularies, Univer mutation/command names the edit journal listens to,
 * and the Home → Cell Styles presets. Extracted from App.tsx verbatim.
 */
import type { CellFormatPatch } from '../domain/workbook-dsl'
import type { WorkbookSnapshot } from '../domain/workbook.types'
import type { ChartEditData } from './WorkbookVisuals'

export const CHART_TYPE_COMMANDS = ['column', 'bar', 'line', 'area', 'pie', 'doughnut'] as const
/// Demo/session charts bake edits into the visual itself — no pending map.
export const EMPTY_CHART_EDITS: ReadonlyMap<string, ChartEditData> = new Map()
export const CHART_LEGEND_COMMANDS = ['right', 'bottom', 'top', 'left', 'none'] as const
export const CHART_LABEL_COMMANDS = ['value', 'percent', 'category-percent', 'none'] as const

/// Ribbon "Change Colors" palettes; applied per series index, cycling.
export const CHART_PALETTES: Record<string, readonly string[]> = {
  office: [
    '#4472c4',
    '#ed7d31',
    '#a5a5a5',
    '#ffc000',
    '#5b9bd5',
    '#70ad47',
    '#264478',
    '#9e480e',
    '#636363',
    '#997300',
  ],
  blue: ['#1f4e79', '#2e75b6', '#5b9bd5', '#9dc3e6', '#bdd7ee', '#2f5597'],
  green: ['#375623', '#548235', '#70ad47', '#a9d18e', '#c5e0b4', '#7fab55'],
  warm: ['#c00000', '#ed7d31', '#ffc000', '#bf8f00', '#843c0c', '#f4b183'],
  gray: ['#404040', '#7f7f7f', '#a6a6a6', '#bfbfbf', '#d0cece', '#595959'],
}

export const initialSnapshot: WorkbookSnapshot = {
  revision: 0,
  sheets: [
    {
      id: 'sheet-1',
      name: 'Sheet1',
      cells: {},
    },
  ],
}

export const FORMULA_MODE_MAX_CELLS = 50_000
export const SET_RANGE_VALUES_MUTATION = 'sheet.mutation.set-range-values'
export const SET_NUMFMT_MUTATION = 'sheet.mutation.set.numfmt'
// Freeze and gridline toggles journal from their mutations so Univer's own
// undo/redo re-records the restored state (ribbon handlers do not record).
export const SET_FROZEN_MUTATION = 'sheet.mutation.set-frozen'
export const TOGGLE_GRIDLINES_MUTATION = 'sheet.mutation.toggle-gridlines'
// Undoing a numfmt set emits the remove mutation; only the ribbon echo cares.
export const REMOVE_NUMFMT_MUTATION = 'sheet.mutation.remove.numfmt'
// Row/column inserts/removals and merges are journaled and replayed at save
// time. Streamed viewports keep working after a shift: read requests and
// results translate through the operation stream (view-transform.ts).
export const STRUCTURAL_EDIT_COMMAND_PATTERN =
  /^sheet\.command\.(insert-row|remove-row|insert-col|remove-col|add-worksheet-merge|remove-worksheet-merge)/
export const ROW_COLUMN_MUTATIONS: Record<
  string,
  { kind: 'insert-rows' | 'remove-rows' | 'insert-cols' | 'remove-cols'; axis: 'row' | 'column' }
> = {
  'sheet.mutation.insert-row': { kind: 'insert-rows', axis: 'row' },
  'sheet.mutation.remove-rows': { kind: 'remove-rows', axis: 'row' },
  'sheet.mutation.insert-col': { kind: 'insert-cols', axis: 'column' },
  'sheet.mutation.remove-col': { kind: 'remove-cols', axis: 'column' },
}
export const MERGE_MUTATIONS: Record<string, 'merge-cells' | 'unmerge-cells'> = {
  'sheet.mutation.add-worksheet-merge': 'merge-cells',
  'sheet.mutation.remove-worksheet-merge': 'unmerge-cells',
}
// Row/column sizing and visibility journal as ordered ops and replay into
// row/cols attributes at save time. (`set-worksheet-row-auto-height` is
// derived layout — it fires during rendering and must NOT journal; the user
// command is `set-worksheet-row-is-auto-height`.)
export const AXIS_ATTR_MUTATIONS: Record<
  string,
  { kind: 'size' | 'hidden' | 'auto-size'; axis: 'row' | 'column'; hidden?: boolean }
> = {
  'sheet.mutation.set-worksheet-row-height': { kind: 'size', axis: 'row' },
  'sheet.mutation.set-worksheet-col-width': { kind: 'size', axis: 'column' },
  'sheet.mutation.set-worksheet-row-is-auto-height': { kind: 'auto-size', axis: 'row' },
  'sheet.mutation.set-row-hidden': { kind: 'hidden', axis: 'row', hidden: true },
  'sheet.mutation.set-row-visible': { kind: 'hidden', axis: 'row', hidden: false },
  'sheet.mutation.set-col-hidden': { kind: 'hidden', axis: 'column', hidden: true },
  'sheet.mutation.set-col-visible': { kind: 'hidden', axis: 'column', hidden: false },
}

/// Univer column pixels → OOXML character width (inverse of
/// characterWidthToPixels), snapped to the format's 1/256 granularity.
export function pixelsToCharacterWidth(pixels: number): number {
  return Math.max(Math.round(((pixels - 5) / 7) * 256) / 256, 1 / 256)
}
// Sorting reorders the model in place; the journal snapshots the sorted
// range afterwards, so the save writes exactly what the screen shows.
export const REORDER_RANGE_MUTATION = 'sheet.mutation.reorder-range'
// Drag-moving a selection follows the same snapshot recipe: journal both the
// vacated source and the landing target as plain cell edits.
export const MOVE_RANGE_MUTATION = 'sheet.mutation.move-range'
export const MOVE_RANGE_COMMAND = 'sheet.command.move-range'
export const CF_RULE_COMMAND_PATTERN = /^sheet\.command\.(add|set)-conditional-rule$/
export const SORT_COMMAND_PATTERN =
  /^sheet\.command\.(sort-range|reorder-range|split-text-to-columns)$/
// Filter changes mark the sheet dirty; the save snapshots the live filter
// model declaratively instead of replaying mutations.
export const FILTER_MUTATIONS = new Set([
  'sheet.mutation.set-filter-criteria',
  'sheet.mutation.set-filter-range',
  'sheet.mutation.remove-filter',
  'sheet.mutation.re-calc-filter',
])
// Conditional-formatting changes mark the sheet dirty; the save snapshots
// the full rule model declaratively (same recipe as filters).
export const CF_MUTATIONS = new Set([
  'sheet.mutation.add-conditional-rule',
  'sheet.mutation.set-conditional-rule',
  'sheet.mutation.delete-conditional-rule',
  'sheet.mutation.move-conditional-rule',
])
// Data-validation changes follow the same declarative recipe as CF: the
// mutation only marks the sheet dirty, the save snapshots the full rule set.
export const DV_MUTATIONS = new Set([
  'data-validation.mutation.addRule',
  'data-validation.mutation.updateRule',
  'data-validation.mutation.removeRule',
])
// Defined-name changes mark a single workbook-level dirty flag; the save
// snapshots the full model (declarative, like CF/DV).
export const DEFINED_NAME_MUTATIONS = new Set([
  'formula.mutation.set-defined-name',
  'formula.mutation.remove-defined-name',
])
// Note edits mark the sheet dirty; the save snapshots the live note set as
// legacy comments. Position-only moves don't change what the file stores.
export const NOTE_MUTATIONS = new Set(['sheet.mutation.update-note', 'sheet.mutation.remove-note'])
// Panel commands that mutate DV rules. Editing needs the sheet's own file
// rules installed first (a declarative save would silently drop them).
export const DV_EDIT_COMMAND_PATTERN =
  /^sheets?\.command\.(addDataValidation|updateDataValidationRuleRange|update-data-validation-(?:setting|options)|clear-range-data-validation|remove-data-validation-rule|remove-all-data-validation)$/
export const FILTER_COMMAND_PATTERN =
  /^sheet\.command\.(set-filter-criteria|set-filter-range|smart-toggle-filter|clear-filter-criteria|remove-sheet-filter|re-calc-filter)$/
// Sheet management is journaled and replayed at save time. Renames rewrite
// qualified references file-side; adds create blank parts; removals fail
// closed when referenced.
export const SHEET_LIFECYCLE_MUTATIONS = new Set([
  'sheet.mutation.insert-sheet',
  'sheet.mutation.remove-sheet',
  'sheet.mutation.set-worksheet-name',
  'sheet.mutation.set-worksheet-order',
  'sheet.mutation.set-worksheet-hidden',
])
// Whole-row moves persist as a journaled move op (the save relocates the
// <row> elements, so heights and hidden flags travel for free); column moves
// are still blocked pending the symmetric <c>/<col> treatment.
export const BLOCKED_COMMAND_PATTERN = /^sheet\.command\.move-cols/
export const MOVE_ROWS_COMMAND = 'sheet.command.move-rows'
// The mutation carries {sourceRange, targetRange} in pre-move coordinates.
export const MOVE_ROWS_MUTATION = 'sheet.mutation.move-rows'
// Sheet duplication clones the worksheet part file-side; the journal records
// the source so the save seeds the new part from it.
export const COPY_SHEET_COMMAND = 'sheet.command.copy-sheet'
// Autofill lands as plain value mutations (journal-friendly), but dragging
// into an unstreamed region would overwrite data the user never saw.
export const AUTO_FILL_COMMAND = 'sheet.command.auto-fill'

/// Home → Cell Styles gallery: Excel's built-in named styles as format
/// patches over the ordinary style pipeline (journal + copy-on-write save).
/// Multiple patches apply in order (per-edge borders need separate calls).
export const CELL_STYLE_PRESETS: Record<string, CellFormatPatch[]> = {
  good: [{ fillColor: '#C6EFCE', fontColor: '#006100' }],
  bad: [{ fillColor: '#FFC7CE', fontColor: '#9C0006' }],
  neutral: [{ fillColor: '#FFEB9C', fontColor: '#9C6500' }],
  input: [
    { fillColor: '#FFCC99', fontColor: '#3F3F76', border: { type: 'all', color: '#7F7F7F' } },
  ],
  output: [
    {
      fillColor: '#F2F2F2',
      fontColor: '#3F3F3F',
      bold: true,
      border: { type: 'all', color: '#3F3F3F' },
    },
  ],
  calculation: [
    {
      fillColor: '#F2F2F2',
      fontColor: '#FA7D00',
      bold: true,
      border: { type: 'all', color: '#7F7F7F' },
    },
  ],
  'warning-text': [{ fontColor: '#FF0000' }],
  title: [{ bold: true, fontSize: 18, fontColor: '#44546A' }],
  'heading-1': [
    {
      bold: true,
      fontSize: 15,
      fontColor: '#44546A',
      border: { type: 'bottom', color: '#4472C4' },
    },
  ],
  'heading-2': [
    {
      bold: true,
      fontSize: 13,
      fontColor: '#44546A',
      border: { type: 'bottom', color: '#8EA9DB' },
    },
  ],
  // Excel's Total is thin-top + double-bottom; the double edge is not in the
  // border command set, so the top rule alone approximates it.
  total: [{ bold: true, border: { type: 'top', color: '#4472C4' } }],
  'accent1-20': [{ fillColor: '#DDEBF7' }],
  'accent1-40': [{ fillColor: '#BDD7EE' }],
  accent1: [{ fillColor: '#4472C4', fontColor: '#FFFFFF' }],
}

/** Legacy localStorage history key: history now lives in project-store (per
 * workbook); this key is only kept for the one-time migration */
export const CHAT_STORAGE_KEY = 'ai-excel-chat-history'

/** Cap on persisted tool args/output in transcripts (the store layer has another
 * 16k truncation as backstop) */
export const PERSIST_TOOL_FIELD_MAX = 16_000

/** Tool args → JSON string (truncated; returns undefined on serialization
 * failure without blocking persistence) */
export function safeJsonInput(input: unknown): string | undefined {
  try {
    const s = JSON.stringify(input)
    return s && s !== '{}' ? s.slice(0, PERSIST_TOOL_FIELD_MAX) : undefined
  } catch {
    return undefined
  }
}
