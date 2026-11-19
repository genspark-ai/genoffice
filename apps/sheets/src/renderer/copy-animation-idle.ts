/**
 * The marching-ants border on a copied range is an animation loop.
 *
 * Univer draws the copy source as a dashed rectangle and moves the dashes one
 * step per frame: the selection control asks for a frame, nudges the dash
 * offset, dirties the scene and asks for the next. The loop runs for as long
 * as the copy mode is on.
 *
 * That is fine for the seconds between Ctrl+C and the paste, and it is not
 * fine for a copy mode that outlives its user. A machine suspended overnight
 * wakes up with the copy mode still on, the loop resumes, and every frame
 * dirties the scene — the frames pile up faster than the collector reclaims
 * them. A 12 KB workbook reached 1.6 GB of renderer memory that way and the
 * app's own watchdog fired (genoffice#2246). The border is a hint for a
 * person who is looking at it; nobody is looking at a window behind another
 * one, or at a border that has been sitting still for a minute.
 *
 * So the loop stops when the document is not on screen and when the reader
 * has left it alone, and starts again when they come back. The copy itself is
 * untouched: the copied range is still there, still pasteable, and the border
 * is still drawn — it just stops moving while nobody is watching.
 */
import { SelectionControl } from '@univerjs/sheets-ui'

/**
 * How long the reader may leave the document alone before the border stops
 * moving. Long enough that a reader who is watching a copy they just made
 * still sees it move, short enough that a forgotten copy does not spend the
 * rest of the session repainting the scene.
 */
export const COPY_ANIMATION_IDLE_MS = 60_000

/** The parts of Univer's selection control this guard touches. */
export interface AntLineControl {
  _startAntLineAnimation: () => void
  _stopAntLineAnimation: () => void
  _disposed?: boolean
}

export interface CopyAnimationIdleOptions {
  /** how long the reader may leave the document alone before the border stops */
  idleMs?: number
  /** whether the document is on screen right now */
  isVisible?: () => boolean
  /** the current time, in ms */
  now?: () => number
}

/**
 * Decides whether a copied range's border may keep moving.
 *
 * The guard is asked once per frame, from the patched animation method, and
 * it answers by letting the frame through or by ending the frame chain — the
 * chain is recursive, so a frame that is not let through is the last one.
 */
export class CopyAnimationIdle {
  private readonly idleMs: number
  private readonly isVisible: () => boolean
  private readonly now: () => number
  /** controls whose border is (or was) moving */
  private readonly tracked = new Set<WeakRef<AntLineControl>>()
  /** controls already known, so a per-frame call does not re-add them */
  private readonly seen = new WeakSet<AntLineControl>()
  /** whether each control's frame chain is alive */
  private readonly running = new WeakMap<AntLineControl, boolean>()
  /** when the reader last touched the document */
  private lastInteraction: number
  /** when the document last went off screen */
  private hiddenAt = 0

  constructor(options: CopyAnimationIdleOptions = {}) {
    this.idleMs = options.idleMs ?? COPY_ANIMATION_IDLE_MS
    this.isVisible = options.isVisible ?? (() => !document.hidden)
    this.now = options.now ?? (() => Date.now())
    this.lastInteraction = this.now()
  }

  /** Whether the border may move right now. */
  get animating(): boolean {
    return this.isVisible() && this.now() - this.lastInteraction < this.idleMs
  }

  /**
   * One frame of the animation.
   *
   * `start` is the original animation step. Not calling it ends the frame
   * chain, which is how the border stops: the animation is recursive, so the
   * step that is not taken is the step that does not happen.
   */
  frame(control: AntLineControl, start: () => void): void {
    if (!this.seen.has(control)) {
      // the border started moving now — that is the moment the reader copied
      // something, so the idle clock starts here rather than at install time
      this.seen.add(control)
      this.tracked.add(new WeakRef(control))
      this.lastInteraction = this.now()
    }
    if (!this.animating) {
      this.running.set(control, false)
      return
    }
    this.running.set(control, true)
    start()
  }

  /**
   * The reader came back: the idle clock starts again and any border that had
   * stopped starts moving.
   *
   * A border that is already moving is left alone — the animation is a
   * recursive chain, and starting a second one would run two animations at
   * once and double what each frame costs.
   */
  resume(): void {
    this.lastInteraction = this.now()
    for (const ref of this.tracked) {
      const control = ref.deref()
      if (!control || control._disposed) {
        this.tracked.delete(ref)
        continue
      }
      if (this.running.get(control)) continue
      // the patched method, so the same rules apply to the frame it opens
      control._startAntLineAnimation()
    }
  }

  /**
   * The document went off screen: stop every border that is moving.
   *
   * The frame chain is cancelled rather than merely left to lapse, because a
   * hidden window may have its frames throttled rather than stopped, and a
   * chain that is still alive would keep the scene dirty.
   */
  pause(): void {
    for (const ref of this.tracked) {
      const control = ref.deref()
      if (!control || control._disposed) {
        this.tracked.delete(ref)
        continue
      }
      this.running.set(control, false)
      control._stopAntLineAnimation()
    }
  }

  /** The document went off screen. */
  away(): void {
    this.hiddenAt = this.now()
    this.pause()
  }

  /**
   * The document came back.
   *
   * A short absence is a glance at another window, and the border picks up
   * where it left off. A long one is a machine that was asleep — a reader who
   * was away that long is not waiting to watch dashes move, and the border
   * stays still until they touch the document again.
   */
  back(): void {
    if (this.now() - this.hiddenAt < this.idleMs) this.resume()
  }
}

/** The reader's own events — any of them means the reader is present. */
const INTERACTION_EVENTS = ['pointerdown', 'pointermove', 'keydown', 'wheel', 'scroll'] as const

let installed = false

/**
 * Installs the guard on Univer's selection control.
 *
 * Returns a function that puts the original animation back.
 */
export function installCopyAnimationIdle(options: CopyAnimationIdleOptions = {}): () => void {
  // Installing twice would wrap the guarded method in itself, and the frames
  // would call each other for ever.
  if (installed) return () => {}
  installed = true

  const guard = new CopyAnimationIdle(options)
  const proto = SelectionControl.prototype as unknown as {
    _startAntLineAnimation: () => void
  }
  const origStart = proto._startAntLineAnimation

  proto._startAntLineAnimation = function (this: AntLineControl): void {
    guard.frame(this, () => origStart.call(this))
  }

  const onInteraction = (): void => {
    guard.resume()
  }
  for (const event of INTERACTION_EVENTS) {
    window.addEventListener(event, onInteraction, { passive: true })
  }
  const onVisibility = (): void => {
    if (document.hidden) guard.away()
    else guard.back()
  }
  document.addEventListener('visibilitychange', onVisibility)

  return () => {
    proto._startAntLineAnimation = origStart
    for (const event of INTERACTION_EVENTS) {
      window.removeEventListener(event, onInteraction)
    }
    document.removeEventListener('visibilitychange', onVisibility)
    installed = false
  }
}
