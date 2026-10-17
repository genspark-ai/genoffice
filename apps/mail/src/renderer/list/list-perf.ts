/**
 * List-side scroll/render perf budgets (H6 owns, P-1..P-5 of the H1
 * hardening spec). Pure helpers plus the search debounce contract so
 * the budgets are enforced as TESTS, not vibes.
 *
 * Seam note (H1 section 4.1/4.2): the debounce delay constant is owned
 * here (H6); the store-query cancellation contract on the main side is
 * owned by H4 (Nadia). On the renderer side a superseded keystroke
 * surfaces as `onSuperseded` (cancel the in-flight store query) and a
 * stale async result is discarded via `SearchQueryTracker.isCurrent`.
 * Synthetic `example.com` mail only; no new prod deps.
 */

/** Default fixed row height in px (mirrors MessageList default). */
export const LIST_ROW_HEIGHT = 76

/** Overscan rows above/below the viewport (mirrors MessageList). */
export const LIST_OVERSCAN = 4

/** Fallback viewport height in px before ResizeObserver measures. */
export const LIST_VIEWPORT_HEIGHT = 480

/** Search-as-you-type debounce floor in ms (P-4: 150-250 window). */
export const SEARCH_DEBOUNCE_MIN_MS = 150

/** Search-as-you-type debounce delay in ms (P-4: single constant). */
export const SEARCH_DEBOUNCE_MS = 200

/** Search-as-you-type debounce ceiling in ms (P-4: 150-250 window). */
export const SEARCH_DEBOUNCE_MAX_MS = 250

export interface ListWindowInput {
  /** Total envelope count (a count, never the array: O(1) by shape). */
  total: number
  scrollTop: number
  viewportHeight: number
  rowHeight: number
  overscan: number
}

export interface ListWindow {
  startIndex: number
  visibleCount: number
  endIndex: number
  topSpacer: number
  bottomSpacer: number
}

/**
 * Virtualized window arithmetic. Pure O(1): only number math, no list
 * traversal, so per-scroll-event cost is independent of mailbox size.
 */
export function computeListWindow(input: ListWindowInput): ListWindow {
  const total = Math.max(0, Math.floor(input.total))
  const rowHeight = Math.max(1, input.rowHeight)
  const overscan = Math.max(0, Math.floor(input.overscan))
  const viewportHeight = Math.max(0, input.viewportHeight)
  const scrollTop = Math.max(0, input.scrollTop)
  const startIndex = Math.min(total, Math.max(0, Math.floor(scrollTop / rowHeight) - overscan))
  const visibleCount = Math.ceil(viewportHeight / rowHeight) + overscan * 2
  const endIndex = Math.min(total, startIndex + visibleCount)
  const topSpacer = startIndex * rowHeight
  const totalHeight = total * rowHeight
  const bottomSpacer = Math.max(0, totalHeight - topSpacer - (endIndex - startIndex) * rowHeight)
  return { startIndex, visibleCount, endIndex, topSpacer, bottomSpacer }
}

/**
 * P-1 bound: maximum mounted rows for a viewport. At spec defaults
 * (480 px viewport, 76 px rows, overscan 4) this is
 * ceil(480/76) + 8 = 15 rows out of 100k envelopes.
 */
export function maxWindowedRows(
  viewportHeight: number,
  rowHeight: number,
  overscan: number,
): number {
  return Math.ceil(Math.max(0, viewportHeight) / Math.max(1, rowHeight)) + Math.max(0, overscan) * 2
}

/** Nearest-rank percentile over sample timings in ms. */
export function percentile(samples: number[], p: number): number {
  if (samples.length === 0) return 0
  const ordered = [...samples].sort((a, b) => a - b)
  const rank = Math.min(ordered.length - 1, Math.max(0, Math.ceil((p / 100) * ordered.length) - 1))
  return ordered[rank] ?? 0
}

/**
 * Debounced search-as-you-type (P-4). `push` on every keystroke:
 * the pending fire is cancelled (superseded callback fires so the
 * caller can cancel the in-flight store query) and the timer
 * restarts. Fires only when the query changed since the last fire,
 * so re-renders with an identical value cost nothing.
 */
export class DebouncedSearch {
  private timer: ReturnType<typeof setTimeout> | null = null
  private latest = ''
  private lastFired: string | null = null
  private generation = 0

  constructor(
    private readonly delayMs: number,
    private readonly onFire: (query: string, generation: number) => void,
    private readonly onSuperseded?: () => void,
  ) {}

  get pending(): boolean {
    return this.timer !== null
  }

  push(query: string): void {
    this.generation += 1
    this.latest = query
    if (this.timer !== null) {
      clearTimeout(this.timer)
      this.timer = null
      this.onSuperseded?.()
    }
    const generation = this.generation
    this.timer = setTimeout(() => {
      this.timer = null
      if (this.latest === this.lastFired) return
      this.lastFired = this.latest
      this.onFire(this.latest, generation)
    }, this.delayMs)
  }

  /** Cancel any pending fire and invalidate the in-flight query. */
  cancel(): void {
    this.generation += 1
    if (this.timer !== null) {
      clearTimeout(this.timer)
      this.timer = null
      this.onSuperseded?.()
    }
  }
}

/**
 * Stale-result guard for async submit paths (P-4): capture `start()`
 * before awaiting the store, then discard the result unless
 * `isCurrent` still holds. A newer keystroke/search always wins;
 * stale results never render.
 */
export class SearchQueryTracker {
  private latest = 0

  start(): number {
    this.latest += 1
    return this.latest
  }

  isCurrent(generation: number): boolean {
    return generation === this.latest
  }
}
