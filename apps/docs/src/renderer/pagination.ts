/**
 * Pagination: greedy page breaking over the continuous-flow render result, by top-level block.
 * Pure functions; all coordinates are content-area Y at 100% zoom (px, 0 = top of page 1 content).
 *
 * Barrel over the pagination-* modules so call sites keep one import path:
 *   pagination-types     block / slice / section geometry records
 *   pagination-sections  section geometry, column specs, doc-grid pitch
 *   pagination-slices    the slicing engine (F2 model) and page queries
 *   pagination-hf        header/footer refs and page-number formatting
 *   pagination-measure   DOM measurement of blocks, notes and float spill
 *   pagination-lines     DOM line-box sampling, line cuts and anchors
 */

export * from './pagination-hf'
export * from './pagination-lines'
export * from './pagination-measure'
export * from './pagination-sections'
export * from './pagination-slices'
export * from './pagination-types'
