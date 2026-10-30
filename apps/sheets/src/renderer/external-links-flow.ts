import type { ExternalLinksUpdateResult } from '../shared/desktop-api'

/** What the links panel's Update action needs from the editor. */
export interface ExternalLinksUpdateDeps {
  /** the workbook path the sidecar reads */
  readonly path: string
  /** the ordinary save; its outcome gates the update */
  readonly save: () => Promise<{ readonly ok: boolean }>
  readonly update: (path: string) => Promise<ExternalLinksUpdateResult>
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
 * A failed first save aborts before the update: the report would describe a
 * file the grid no longer matches.
 */
export async function updateExternalLinksFlow(deps: ExternalLinksUpdateDeps): Promise<boolean> {
  if (!(await deps.save()).ok) return false
  const report = await deps.update(deps.path)
  deps.apply(report)
  deps.onStatuses(report)
  await deps.save()
  return true
}
