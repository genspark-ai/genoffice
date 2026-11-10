/**
 * Pure logic for slide shows: playback sequence computation + rehearsal timing accumulation.
 * Extracted from SlideShowView for unit testing (no React/DOM dependency).
 */

/** Custom show: subset of slides in user-specified order (original indexes). App persists it per document to localStorage. */
export interface CustomShow {
  id: string
  name: string
  slideIndices: number[]
}

/**
 * Compute the playback sequence (array of original indexes).
 * - Default: all slides in order, skipping hidden ones (starting from a hidden slide still plays it)
 * - Non-empty customOrder: play in its order (out-of-range slides filtered; hidden slides still skipped, except the start slide)
 * - Fallback: when the result is empty, play the start slide; an out-of-range start yields []
 */
export function computePlayOrder(
  slides: ReadonlyArray<{ hidden?: boolean }>,
  startAt: number,
  customOrder?: readonly number[],
): number[] {
  // A stale startAt (e.g. after slide deletions) must not emit an
  // out-of-range slide: nothing is playable from an invalid start.
  if (!Number.isInteger(startAt) || startAt < 0 || startAt >= slides.length) return []
  const playable = (i: number) => slides[i] != null && (!slides[i]!.hidden || i === startAt)
  const o =
    customOrder && customOrder.length > 0
      ? customOrder.filter(playable)
      : slides.map((_, i) => i).filter(playable)
  return o.length > 0 ? [...o] : [startAt]
}

// ── Rehearsal timing ─────────────────────────────────────────────────────────────

/** Rehearsal timing state: perPageMs accumulates dwell milliseconds by original slide index. */
export interface RehearseTiming {
  perPageMs: number[]
  /** Slide currently dwelt on (original index; -1 = finished) */
  currentIndex: number
  /** Timestamp of entering the current slide (ms) */
  enteredAt: number
}

/** Start rehearsal: begin timing from startIndex. */
export function startRehearse(slideCount: number, startIndex: number, now: number): RehearseTiming {
  return {
    perPageMs: new Array(Math.max(0, slideCount)).fill(0),
    currentIndex: startIndex,
    enteredAt: now,
  }
}

/** Page turn: accumulate the current slide's dwell into perPageMs, then switch to nextIndex and restart timing (revisiting a slide keeps accumulating). */
export function switchRehearsePage(
  t: RehearseTiming,
  nextIndex: number,
  now: number,
): RehearseTiming {
  const perPageMs = t.perPageMs.slice()
  if (t.currentIndex >= 0 && t.currentIndex < perPageMs.length) {
    perPageMs[t.currentIndex]! += Math.max(0, now - t.enteredAt)
  }
  return { perPageMs, currentIndex: nextIndex, enteredAt: now }
}

/** End rehearsal: accumulate the last slide's dwell, then convert to seconds per slide (rounded; visited slides count at least 1 second). */
export function finishRehearse(t: RehearseTiming, now: number): number[] {
  const final = switchRehearsePage(t, -1, now)
  return final.perPageMs.map((ms) => (ms > 0 ? Math.max(1, Math.round(ms / 1000)) : 0))
}

/** m:ss clock display (rehearsal timer bar / save confirmation dialog). */
export function formatClock(ms: number): string {
  const sec = Math.max(0, Math.floor(ms / 1000))
  return `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, '0')}`
}
