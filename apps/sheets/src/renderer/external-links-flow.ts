import type {
  ExternalLinkOps,
  ExternalLinkUsageCell,
  ExternalLinkUsageResult,
  ExternalLinksUpdateResult,
} from '../shared/desktop-api'

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

/** What the links panel's Break action needs from the editor. */
export interface ExternalLinkBreakDeps {
  /** the 1-based `[n]` link to break */
  readonly index: number
  /** the workbook's referencing formula cells */
  readonly readUsage: () => Promise<ExternalLinkUsageResult>
  /**
   * Rewrites the link's formula cells as the literals their caches hold and
   * resolves to the undo for that write.
   */
  readonly journal: (cells: readonly ExternalLinkUsageCell[]) => Promise<() => Promise<void>>
  /** hands the ops to the next save */
  readonly queue: (ops: ExternalLinkOps) => void
  readonly clear: () => void
  readonly save: () => Promise<{ readonly ok: boolean; readonly reason?: string }>
}

/**
 * The panel's Break action: journal the link's cells as literals (the grid's
 * current values beat the file's possibly stale caches), queue the break for
 * one save, then roll the formulas back when that save does not land — a grid
 * showing literals for a file that still has formulas and the link is exactly
 * the inconsistency the journal-before-save order used to leave behind.
 */
export async function breakExternalLinkFlow(deps: ExternalLinkBreakDeps): Promise<boolean> {
  const usage = await deps.readUsage()
  const entry = usage.links.find((link) => link.index === deps.index)
  const restore = entry === undefined ? null : await deps.journal(entry.cells)
  deps.queue({ broken: [deps.index], retargets: [] })
  try {
    const outcome = await deps.save()
    if (!outcome.ok && restore !== null) await restore()
    return outcome.ok
  } finally {
    deps.clear()
  }
}
