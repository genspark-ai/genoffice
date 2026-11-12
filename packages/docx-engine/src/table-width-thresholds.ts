/**
 * Tolerances of the tcW-vs-tblGrid normalisation in parse.ts (extractTableModel).
 * They decide whether a saved grid is Word's own layout or a generator placeholder
 * that yields to w:tcW / w:tblW. None has a Word probe: each is corpus-fitted on
 * the sample that motivated it, so the renderer's table-layout/ lists them with
 * that provenance and a probe may replace any of them.
 */

/** every gridCol within this share of the mean (floor UNIFORM_GRID_MIN_TWIPS) is an evenly
 * split placeholder, not a laid-out grid — corpus-fitted 2026-09-17, unprobed */
export const UNIFORM_GRID_TOLERANCE = 0.01
export const UNIFORM_GRID_MIN_TWIPS = 2

/** a uniform grid whose column shares differ from tcW by more than this many percentage
 * points is stale — corpus-fitted 2026-08-07, unprobed */
export const TCW_GRID_PCT_TOLERANCE = 2

/** Σ grid may differ from Σ tcW (or from a dxa tblW) by this many twips per column before
 * the grid counts as a different layout: per-column rounding — corpus-fitted 2026-08-31 /
 * 2026-09-17, unprobed */
export const GRID_ROUNDING_TWIPS_PER_COLUMN = 1

/** with w:tblCellSpacing the grid also holds the gaps tcW leaves out: Word saves 2s per
 * column plus s on the outer two, other generators 3s each — corpus-fitted 2026-09-24, unprobed */
export const CELL_SPACING_GAPS_PER_COLUMN = 3

/** Σ grid within this share of the text column (tblW auto shrunk to it) or of the pct width
 * is Word's finished layout — corpus-fitted 2026-09-17; the pct case probed 2026-10-29 */
export const SAVED_GRID_TOLERANCE = 0.03
