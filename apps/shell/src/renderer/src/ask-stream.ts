/**
 * The Ask panel's streaming render discipline, extracted so it is testable
 * without a DOM: an animation-frame flush queue that coalesces a token burst
 * into one React render, and the stick-to-bottom decision the thread scroll
 * follows.
 *
 * onText can fire dozens of times between two paints, and re-rendering the
 * whole thread per token — every history row re-created, every Markdown
 * re-parsed — made the streaming animation O(history × tokens) on the one
 * surface every summon path lands on (global hotkey, tray, deep link, Ctrl+K).
 * The queue defers only the React flush; the panel's ref is still written
 * synchronously, so persistence and finish logic see the final text before
 * React has painted it.
 */

/** how a flush is scheduled; injectable so tests drive frames by hand */
export type ScheduleFlush = (run: () => void) => unknown

/** requestAnimationFrame where present, a macrotask where it is not */
export const scheduleRaf: ScheduleFlush = (run) =>
  typeof requestAnimationFrame === 'function'
    ? requestAnimationFrame(run)
    : window.setTimeout(run, 0)

/**
 * At most one flush per animation frame. Requests made while one is pending
 * fold into it: the flush reads whatever the caller's ref holds when the frame
 * lands, not a snapshot taken at request time, so every patch scheduled in
 * between rides the same frame. A synchronous flush elsewhere leaves the
 * pending frame setting the same array reference, which React's Object.is
 * bailout turns into a no-op render.
 */
export class RafFlushQueue {
  private scheduled = false
  private disposed = false
  constructor(
    private readonly flush: () => void,
    private readonly schedule: ScheduleFlush = scheduleRaf,
  ) {}

  /** ask for a flush on the next frame; requests while one is pending fold in */
  request(): void {
    if (this.disposed || this.scheduled) return
    this.scheduled = true
    this.schedule(() => {
      this.scheduled = false
      if (!this.disposed) this.flush()
    })
  }

  /** stop flushing (unmount); a frame already in flight lands as a no-op */
  dispose(): void {
    this.disposed = true
  }
}

/** the thread counts as pinned to the bottom within this many pixels */
export const STICK_TO_BOTTOM_PX = 48

/**
 * Stick-to-bottom decision for the thread's scroll guard: the view follows
 * the conversation's growth only while it sits this close to the bottom, so
 * scrolling up to reread an earlier answer is never yanked back down
 * mid-stream (the same guard the editors' AiPanel uses).
 */
export function sticksToBottom(
  scrollTop: number,
  scrollHeight: number,
  clientHeight: number,
  threshold: number = STICK_TO_BOTTOM_PX,
): boolean {
  return scrollHeight - scrollTop - clientHeight < threshold
}
