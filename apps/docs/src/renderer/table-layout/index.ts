/**
 * Table column widths: one algorithm from the parsed TableModel (tblW / tblLayout /
 * tcW / gridCol already normalised by docx-engine) to the grid the renderer draws.
 * Every width decision after parse lives here; convert.ts / table-geometry.ts
 * only call in. Rules enter with a dated Word probe (word-truths.md) and a node
 * unit test (apps/docs/tests/table-layout.test.ts); tools/check-table-layout-rules.mjs
 * keeps new heuristics out of the callers.
 */
export {
  layoutTableColumns,
  clampColumnWidths,
  expandAutofitColWidths,
  clampTableColWidths,
  type ColumnBudget,
  type ColumnLayoutOptions,
  type ColumnPhase,
  type ColumnWidths,
} from './column-widths'
export {
  distributeContentAutofit,
  reclaimFloorTwips,
  nestedFloorTwips,
  reclaimSurplus,
  scaleToFit,
} from './distribute'
export * from './constants'
export { gridFromEditorColumns } from './editor-grid'
export {
  cellMeasurePx,
  cellSidePadTwips,
  maxContentColTwips,
  minContentColTwips,
  nestedTableMinTwips,
  nestedTablePreferredTwips,
  sideMarginTwips,
  tableIndentTwips,
} from './content-widths'
export {
  textMeasureOf,
  defaultTextMeasure,
  type MeasureRun,
  type RunDefaults,
  type TextMeasure,
} from './measure'
export { maxLineWidthPx, maxWordWidthPx, middleDotGlue, type MiddleDotGlue } from './tokenize'
export { budgetAvail, sectionWidthBudget, type TableWidthBudget } from './section-budget'
