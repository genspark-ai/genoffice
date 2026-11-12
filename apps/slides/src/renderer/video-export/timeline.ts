/**
 * Pure schedule of an exported show: when each slide enters, how long its transition runs,
 * when each animation step starts, and how long the slide stays. One frame at absolute time
 * t maps back to (slide, transition progress | animation cursor) through this table.
 */
import type { AnimStep } from '../animation-play'
import type { TransitionKind } from '../../shared/ipc'

/** Pause between two animation steps that would each wait for a click in a live show. */
export const STEP_GAP_MS = 500
/** Time the finished animation state stays on screen before the slide may end. */
export const ANIM_HOLD_MS = 1000

export type PlayedTransition = Exclude<TransitionKind, 'none' | 'random'>

/** Duration of each transition, matching the show's CSS approximations and the Morph tween. */
export const TRANSITION_MS: Record<PlayedTransition, number> = {
  morph: 500,
  fade: 500,
  push: 500,
  wipe: 600,
  split: 550,
  circle: 600,
  cover: 500,
  pull: 500,
  dissolve: 600,
  zoom: 500,
}

const RANDOM_POOL: PlayedTransition[] = [
  'fade',
  'push',
  'wipe',
  'split',
  'circle',
  'cover',
  'pull',
  'dissolve',
  'zoom',
]

export interface SlideInput {
  /** Index into the deck (hidden slides are not passed) */
  index: number
  transition: TransitionKind
  /** Saved advance time (rehearsal / recording), when any */
  advanceMs: number | null
  /** Total length of the slide's recorded narration */
  narrationMs: number
  steps: AnimStep[]
}

export interface TimelineOptions {
  /** Time on a slide that has no saved advance time (or when saved timings are ignored) */
  defaultSlideMs: number
  /** Honor saved advance times and keep the slide up for its narration */
  useTimings: boolean
}

export interface SlideSegment {
  index: number
  /** Absolute start; the transition into the slide begins here */
  startMs: number
  /** Whole segment, transition included */
  durationMs: number
  transition: PlayedTransition | null
  transitionMs: number
  steps: AnimStep[]
  /** Start of each step, relative to the segment start */
  stepStartsMs: number[]
}

export interface AnimCursor {
  played: number
  /** Time into the step being played; null when resting between steps or after the last one */
  activeMs: number | null
}

/** Resolve 'random' deterministically so a re-export of the same deck is the same video. */
export function resolveTransition(kind: TransitionKind, slot: number): PlayedTransition | null {
  if (kind === 'none') return null
  if (kind === 'random') return RANDOM_POOL[slot % RANDOM_POOL.length]!
  return kind
}

export function buildTimeline(slides: SlideInput[], opts: TimelineOptions): SlideSegment[] {
  const out: SlideSegment[] = []
  let cursor = 0
  slides.forEach((s, slot) => {
    const transition = resolveTransition(s.transition, slot)
    const transitionMs = transition ? TRANSITION_MS[transition] : 0
    const stepStartsMs: number[] = []
    let at = transitionMs
    s.steps.forEach((step, i) => {
      if (!(i === 0 && step.auto)) at += STEP_GAP_MS
      stepStartsMs.push(at)
      at += step.totalMs
    })
    const animEndMs = s.steps.length ? at + ANIM_HOLD_MS : 0
    const saved = opts.useTimings ? s.advanceMs : null
    const dwellMs = saved ?? opts.defaultSlideMs
    const narrationMs = opts.useTimings ? s.narrationMs : 0
    const durationMs = Math.max(transitionMs + dwellMs, animEndMs, transitionMs + narrationMs)
    out.push({
      index: s.index,
      startMs: cursor,
      durationMs,
      transition,
      transitionMs,
      steps: s.steps,
      stepStartsMs,
    })
    cursor += durationMs
  })
  return out
}

export function totalDurationMs(segments: SlideSegment[]): number {
  const last = segments[segments.length - 1]
  return last ? last.startMs + last.durationMs : 0
}

/** The segment on screen at absolute time t (the last one past the end). */
export function segmentAt(segments: SlideSegment[], tMs: number): SlideSegment | null {
  let lo = 0
  let hi = segments.length - 1
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1
    if (segments[mid]!.startMs <= tMs) lo = mid
    else hi = mid - 1
  }
  return segments[lo] ?? null
}

/** Animation cursor of a segment at time t relative to the segment start. */
export function animCursorAt(seg: SlideSegment, tRelMs: number): AnimCursor {
  let played = 0
  for (let i = 0; i < seg.steps.length; i++) {
    const start = seg.stepStartsMs[i]!
    const end = start + seg.steps[i]!.totalMs
    if (tRelMs >= end) played++
    else if (tRelMs >= start) return { played, activeMs: tRelMs - start }
    else break
  }
  return { played, activeMs: null }
}
