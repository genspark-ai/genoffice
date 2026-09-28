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

export function installStatusBarStatsFilter(
  runtime: UniverRuntime,
  getEnabled: () => readonly StatusBarFunc[],
): StatusBarStatsFilter {
  const service = runtime.univer.__getInjector().get(IStatusBarService) as {
    setState(state: StatusBarState | null): void
  }
  const upstreamSetState = service.setState.bind(service)
  let lastState: StatusBarState | null = null
  service.setState = (state) => {
    lastState = state
    upstreamSetState(filterStatusBarState(state, getEnabled()))
  }
  return {
    refresh() {
      upstreamSetState(filterStatusBarState(lastState, getEnabled()))
    },
    dispose() {
      service.setState = upstreamSetState
    },
  }
}
