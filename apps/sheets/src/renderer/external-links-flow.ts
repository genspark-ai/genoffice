import type { ExternalLinksUpdateResult } from '../shared/desktop-api'

/** What the links panel's Update action needs from the editor. */
export interface ExternalLinksUpdateDeps {
  /**
   * The ordinary save. `reason: 'no-edits'` is a benign no-op (nothing was
   * pending, so the disk already matches the grid); any other failure aborts.
   */
  readonly save: () => Promise<{ readonly ok: boolean; readonly reason?: string }>
  /** Re-reads the reachable sources of the session's workbook. */
  readonly update: () => Promise<ExternalLinksUpdateResult>
  /** installs the report's values into the recalc overlay */
  readonly apply: (report: ExternalLinksUpdateResult) => void
  /** the report's links are the first reachability verdict */
  readonly onStatuses: (report: ExternalLinksUpdateResult) => void
}

/**
 * The panel's Update action, in the only order that keeps the live grid and
 * the file on disk in step: save first (the sidecar reads the file on disk and
 * the values are applied by coordinate to the grid, so an unsaved row/column
 * insert would shift them), update, apply the overlay, then save again so the
 * applied values become the formulas' cached `<v>`.
 *
 * The first save is allowed to be a no-op: with nothing edited it reports
 * `no-edits` and the disk already matches the grid. A failed first save aborts
 * before the update: the report would describe a file the grid no longer
 * matches.
 */
export async function updateExternalLinksFlow(deps: ExternalLinksUpdateDeps): Promise<boolean> {
  const saved = await deps.save()
  if (!saved.ok && saved.reason !== 'no-edits') return false
  const report = await deps.update()
  deps.apply(report)
  deps.onStatuses(report)
  await deps.save()
  return true
}
