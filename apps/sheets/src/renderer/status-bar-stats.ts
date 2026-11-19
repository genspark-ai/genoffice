/**
 * Excel-style status bar statistics preferences: which of the six aggregates
 * the footer shows. Univer's IStatusBarService has no visibility filter, so
 * the installed hook wraps its setState and drops unticked functions before
 * the React footer sees them (an item missing from the values list renders
 * as disabled, i.e. hidden). Click-to-copy and the streamed-file recompute
 * are untouched: both go through the same wrapped setState.
 */
import { IStatusBarService } from '@univerjs/sheets-ui'

import type { UniverRuntime } from './univer-state'

export const STATUS_BAR_FUNCS = ['AVERAGE', 'COUNTA', 'COUNT', 'MIN', 'MAX', 'SUM'] as const
export type StatusBarFunc = (typeof STATUS_BAR_FUNCS)[number]

/** Excel's out-of-the-box ticks: Average, Count, Sum. */
export const DEFAULT_STATUS_BAR_FUNCS: readonly StatusBarFunc[] = ['AVERAGE', 'COUNTA', 'SUM']

export const STATUS_BAR_STATS_STORAGE_KEY = 'sheets.statusBarStats'

interface StorageLike {
  getItem(key: string): string | null
  setItem(key: string, value: string): void
}

export function readStatusBarFuncs(storage: StorageLike | null | undefined): StatusBarFunc[] {
  try {
    const raw = storage?.getItem(STATUS_BAR_STATS_STORAGE_KEY)
    if (raw === null || raw === undefined) return [...DEFAULT_STATUS_BAR_FUNCS]
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return [...DEFAULT_STATUS_BAR_FUNCS]
    return STATUS_BAR_FUNCS.filter((func) => parsed.includes(func))
  } catch {
    return [...DEFAULT_STATUS_BAR_FUNCS]
  }
}

export function writeStatusBarFuncs(
  storage: StorageLike | null | undefined,
  funcs: readonly StatusBarFunc[],
): void {
  try {
    storage?.setItem(STATUS_BAR_STATS_STORAGE_KEY, JSON.stringify(funcs))
  } catch {
    // Storage may be full or disabled; the in-memory choice still applies.
  }
}

export function toggleStatusBarFunc(
  funcs: readonly StatusBarFunc[],
  func: StatusBarFunc,
): StatusBarFunc[] {
  const next = funcs.includes(func) ? funcs.filter((item) => item !== func) : [...funcs, func]
  return STATUS_BAR_FUNCS.filter((item) => next.includes(item))
}

interface StatusBarValue {
  func: string
  value: number
}

interface StatusBarState {
  values: StatusBarValue[]
  pattern: string | null
}

export function filterStatusBarState(
  state: StatusBarState | null,
  enabled: readonly StatusBarFunc[],
): StatusBarState | null {
  if (!state) return null
  return {
    values: state.values.filter((item) => (enabled as readonly string[]).includes(item.func)),
    pattern: state.pattern,
  }
}

export interface StatusBarStatsFilter {
  /** Re-applies the last upstream state after the preference changes. */
  refresh(): void
  dispose(): void
}

interface StatusBarServiceFunction {
  func: string
  filter?: ((state: StatusBarState) => boolean) | undefined
}

/**
 * Univer's StatusBarService decides per-function visibility (SUM/AVERAGE/
 * MIN/MAX need COUNTA > 1 and COUNT > 0) by reading COUNT/COUNTA out of the
 * values list it is given. Those rules must run on the FULL upstream state:
 * once the preference filter removes an unticked COUNT entry, the service
 * reads COUNT as 0 and hides every numeric statistic (the footer kept only
 * the COUNTA count for any selection). Run the service's own rules here, then hand it a
 * pre-decided list with the rules neutralized.
 */
export function applyStatusBarVisibilityRules(
  state: StatusBarState | null,
  rules: ReadonlyMap<string, ((state: StatusBarState) => boolean) | undefined>,
): StatusBarState | null {
  if (!state) return null
  return {
    values: state.values.filter((item) => {
      const rule = rules.get(item.func)
      return rule === undefined || rule(state)
    }),
    pattern: state.pattern,
  }
}

export function installStatusBarStatsFilter(
  runtime: UniverRuntime,
  getEnabled: () => readonly StatusBarFunc[],
): StatusBarStatsFilter {
  // Univer's status types are not exported; the service's `pattern` is a
  // wider Nullable, so the cast has to go through unknown.
  const service = runtime.univer.__getInjector().get(IStatusBarService) as unknown as {
    setState(state: StatusBarState | null): void
    getFunctions?(): StatusBarServiceFunction[]
  }
  const upstreamSetState = service.setState.bind(service)
  const entries = service.getFunctions?.() ?? []
  const rules = new Map(entries.map((entry) => [entry.func, entry.filter]))
  for (const entry of entries) entry.filter = undefined
  let lastState: StatusBarState | null = null
  const publish = (state: StatusBarState | null): void => {
    upstreamSetState(
      filterStatusBarState(applyStatusBarVisibilityRules(state, rules), getEnabled()),
    )
  }
  service.setState = (state) => {
    lastState = state
    publish(state)
  }
  return {
    refresh() {
      publish(lastState)
    },
    dispose() {
      service.setState = upstreamSetState
      for (const entry of entries) entry.filter = rules.get(entry.func)
    },
  }
}
