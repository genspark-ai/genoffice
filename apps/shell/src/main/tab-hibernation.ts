/**
 * Which background tabs to put to sleep. Pure: the tab manager feeds it a
 * snapshot and acts on the ids it returns.
 *
 * Every live editor tab is its own renderer process plus its share of the GPU
 * process, and neither shrinks while the tab is merely hidden (measured:
 * about 250 MB per Docs tab, released only when the webContents is
 * destroyed). So a tab that has not been looked at for a while is destroyed
 * and rebuilt from its file on return — but only once there are more live
 * tabs than the machine comfortably holds, so a handful of tabs never pays
 * the reload.
 */

export interface SleepCandidate {
  id: string
  /** when the tab was last the active one (ms epoch) */
  lastActiveAt: number
}

export interface SleepPolicy {
  /** live tabs kept regardless of idleness; beyond this, idle tabs sleep */
  maxLiveTabs: number
  /** a tab must have been in the background at least this long */
  idleMs: number
}

export const DEFAULT_IDLE_MS = 10 * 60_000
export const SLEEP_SWEEP_INTERVAL_MS = 60_000
/** share of physical memory the app's processes may use before idleness stops mattering */
export const MEMORY_BUDGET_SHARE = 0.3

/** 4 live tabs on an 8 GB machine, 8 on 16 GB, 12 from 32 GB up */
export function defaultMaxLiveTabs(totalMemoryBytes: number): number {
  const gb = totalMemoryBytes / 2 ** 30
  if (gb <= 8) return 4
  if (gb <= 16) return 8
  return 12
}

/**
 * Oldest-first ids to sleep: candidates beyond `maxLiveTabs` that have been
 * idle long enough. `liveCount` is every live tab (including the active one
 * and non-candidates such as pinned or dirty tabs) — the cap is about process
 * count, not about who is eligible. Over the memory budget neither the cap
 * nor idleness matters: the single oldest candidate sleeps per sweep, so a
 * few large documents are reclaimed gradually rather than all at once.
 */
export function pickTabsToSleep(
  candidates: readonly SleepCandidate[],
  liveCount: number,
  policy: SleepPolicy,
  now: number,
  overBudget = false,
): string[] {
  const byAge = [...candidates].sort((a, b) => a.lastActiveAt - b.lastActiveAt)
  if (overBudget) return byAge.slice(0, 1).map((c) => c.id)
  const excess = liveCount - policy.maxLiveTabs
  if (excess <= 0) return []
  return byAge
    .filter((c) => now - c.lastActiveAt >= policy.idleMs)
    .slice(0, excess)
    .map((c) => c.id)
}
