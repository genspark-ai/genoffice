import type { BlockBox } from './pagination-types'

/**
 * Lookups the pagination pass repeats once per page over the block list.
 * Linear scans made those blocks × pages (a 10k-block, 650-page document ran
 * six million comparisons per fixed-point iteration); the sorted copies here
 * answer each in log time with the same first-in-document-order result.
 */

/** ascending numbers with tolerance queries */
export class SortedYs {
  private readonly ys: number[]

  constructor(ys: Iterable<number>) {
    this.ys = [...new Set(ys)].sort((a, b) => a - b)
  }

  get length(): number {
    return this.ys.length
  }

  /** index of the first entry >= y */
  lowerBound(y: number): number {
    let lo = 0
    let hi = this.ys.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (this.ys[mid] < y) lo = mid + 1
      else hi = mid
    }
    return lo
  }

  /** some entry strictly inside (a, b) */
  hasBetween(a: number, b: number): boolean {
    const i = this.lowerBound(a)
    const j = this.ys[i] === a ? i + 1 : i
    return j < this.ys.length && this.ys[j] < b
  }

  /** some entry strictly within tol of y (|entry - y| < tol) */
  hasNear(y: number, tol: number): boolean {
    let i = this.lowerBound(y - tol)
    if (i < this.ys.length && this.ys[i] === y - tol) i++
    return i < this.ys.length && this.ys[i] < y + tol
  }

  /** last entry <= y, or undefined */
  floor(y: number): number | undefined {
    const i = this.lowerBound(y + Number.EPSILON)
    return i > 0 ? this.ys[i - 1] : undefined
  }
}

/** Blocks by top: the first block in document order at or containing a flow Y. */
export class BlockIndex {
  private readonly order: number[]
  private readonly tops: number[]
  private readonly maxHeight: number

  constructor(private readonly blocks: BlockBox[]) {
    this.order = blocks.map((_, i) => i).sort((a, b) => blocks[a].top - blocks[b].top || a - b)
    this.tops = this.order.map((i) => blocks[i].top)
    this.maxHeight = blocks.reduce((m, b) => Math.max(m, b.height), 0)
  }

  private lowerBound(y: number): number {
    let lo = 0
    let hi = this.tops.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (this.tops[mid] < y) lo = mid + 1
      else hi = mid
    }
    return lo
  }

  /** index of the first block whose top is within 0.5 px of y (and passes pred), or -1 */
  firstAtTop(y: number, pred?: (b: BlockBox) => boolean): number {
    let best = -1
    for (let k = this.lowerBound(y - 0.5); k < this.tops.length && this.tops[k] < y + 0.5; k++) {
      const i = this.order[k]
      if (Math.abs(this.blocks[i].top - y) < 0.5 && (!pred || pred(this.blocks[i]))) {
        if (best < 0 || i < best) best = i
      }
    }
    return best
  }

  /** first block in document order with top < y < top + height - 0.5 (and passing pred) */
  containing(y: number, pred?: (b: BlockBox) => boolean): BlockBox | undefined {
    let best = -1
    const end = this.lowerBound(y)
    for (let k = this.lowerBound(y - this.maxHeight); k < end; k++) {
      const i = this.order[k]
      const b = this.blocks[i]
      if (b.top < y && y < b.top + b.height - 0.5 && (!pred || pred(b))) {
        if (best < 0 || i < best) best = i
      }
    }
    return best >= 0 ? this.blocks[best] : undefined
  }
}

/**
 * Capacity windows (column ranges of mixed-column regions, titlePg first
 * pages) by start, duplicates kept: the smallest window capacity containing a
 * Y. A running max end bounds the back scan (windows are near-disjoint, so it
 * visits the few that can still contain the Y).
 */
export class CapacityWindows {
  private readonly windows: Array<{ start: number; end: number; cap: number }>
  private readonly maxEnd: number[] = []

  constructor(windows: Iterable<{ start: number; end: number; cap: number }>) {
    this.windows = [...windows].sort((a, b) => a.start - b.start)
    let m = -Infinity
    for (const w of this.windows) this.maxEnd.push((m = Math.max(m, w.end)))
  }

  /** min(base, cap of every window with start - 0.5 <= y < end - 0.5) */
  capAt(y: number, base: number): number {
    // first window starting past the tolerance; the scan walks back from there
    let lo = 0
    let hi = this.windows.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (this.windows[mid].start <= y + 0.5) lo = mid + 1
      else hi = mid
    }
    let cap = base
    for (let k = lo - 1; k >= 0; k--) {
      if (this.maxEnd[k] - 0.5 <= y) break
      const w = this.windows[k]
      if (w.start - 0.5 <= y && y < w.end - 0.5) cap = Math.min(cap, w.cap)
    }
    return cap
  }
}
