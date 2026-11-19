/**
 * Every numeric threshold of the column-width algorithm, with its provenance. A value
 * marked "unprobed" is corpus-fitted and may be replaced by a Word probe; the parse-side
 * tolerances (tcW vs tblGrid normalisation) live in docx-engine and are listed here.
 */
export {
  CELL_SPACING_GAPS_PER_COLUMN,
  GRID_ROUNDING_TWIPS_PER_COLUMN,
  SAVED_GRID_TOLERANCE,
  TCW_GRID_PCT_TOLERANCE,
  UNIFORM_GRID_MIN_TWIPS,
  UNIFORM_GRID_TOLERANCE,
} from '@genoffice/docx-engine'

/** w:tblCellMar default of TableNormal (ECMA-376; probe 2026-11-09: 30 m -> 30 x 175.7 + 216) */
export const DEFAULT_CELL_MAR = 108

/**
 * Heuristic-measure slack: browser fallback faces run wider than the 0.52 em average
 * (Arial / Helvetica digits and most lowercase 0.556 em, +7 %); an under-floored column
 * re-shatters the very word the floor exists for (RFP sample: two-digit row numbers broke
 * one digit per line, Word grants ~11 % over the estimate). Measured widths get no slack
 * (probe: 181.9 pt measured vs 181.7 drawn) — corpus-fitted 2026-08, unprobed.
 */
export const MIN_CONTENT_SLACK = 1.08

/** compression floor for real columns (24 px); sliver grid slots below it keep their
 * width — corpus-fitted 2026-08, unprobed */
export const MIN_COL_TWIPS = 360
