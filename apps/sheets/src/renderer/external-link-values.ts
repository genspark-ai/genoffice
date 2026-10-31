import type { ExternalLinkCell } from '../shared/desktop-api'
import type { PinnedClosureCell } from './univer-state'

/**
 * The recalculation-overlay cell for one entry of the external-links update
 * report, or null when the report carries no fresh value for that cell (the
 * file never cached one): the cell keeps whatever the file has.
 *
 * The report's `kind` decides the type the next save writes into the formula's
 * cached `<v>`: a boolean result becomes a boolean overlay value (saved as
 * `t="b"`), an error becomes an overlay value flagged `isError` (saved as
 * `t="e"`), text stays text. Without the kind, a TRUE result would land in the
 * file as the string "TRUE".
 */
export function externalLinkOverlayCell(cell: ExternalLinkCell): PinnedClosureCell | null {
  switch (cell.kind) {
    case 'number': {
      const numeric = cell.number ?? Number(cell.formatted)
      // a number the report cannot spell is not a value worth caching
      return Number.isFinite(numeric) ? { v: numeric } : null
    }
    case 'boolean':
      return { v: /^(?:true|1)$/i.test(cell.formatted) }
    case 'error':
      return { v: cell.formatted, isError: true }
    case 'text':
      return { v: cell.formatted }
    default:
      return null
  }
}
