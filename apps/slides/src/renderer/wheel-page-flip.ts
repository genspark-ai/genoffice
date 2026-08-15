// Wheel/trackpad paging for the editing stage (PowerPoint/WPS behavior: while
// the slide fits the viewport, plain scrolling turns pages).
//
// The hard part is telling one user intent apart from a stream of WheelEvents:
//  - a mouse wheel emits one large delta per notch, with real time between
//    notches → flip per notch;
//  - a trackpad swipe is a continuous ~8-16ms stream: it ramps up from small
//    deltas (or opens notch-sized when flicked hard / frame-coalesced) and then
//    coasts on a decaying momentum tail that itself contains notch-sized
//    deltas → the whole gesture must produce exactly one flip.
// Chromium exposes no momentum phase, so intent is reconstructed from the
// stream itself, with three independent signals:
//  - gesture density: wheel gestures are sparse (nearly every event arrives
//    ≥ DISCRETE_MS after the previous one), trackpad gestures are dense. A
//    notch-sized delta flips per-notch only while its gesture stays sparse —
//    one rapid interloper is tolerated (wheel jitter), a second reclassifies
//    the gesture as touch and swallows the rest;
//  - a quiet gap or a sign change starts a new gesture;
//  - once a gesture has flipped and is being swallowed, only a *rising per-ms
//    velocity* (a fresh touch — real momentum only ever decays) gets events
//    considered again. Spacing alone can't reopen it: the flip itself
//    re-renders the slide, and that jank coalesces the very next momentum
//    frames into notch-sized deltas spaced past DISCRETE_MS. Velocity is
//    immune to coalescing (delta and interval grow together).

/** A single event at/above this is a notch; smaller deltas accumulate toward
 * it before the first flip of a trackpad gesture. */
const NOTCH = 60
/** Silence longer than this ends the current gesture (momentum events arrive
 * well under this apart, deliberate consecutive swipes well over it). */
const GAP_MS = 200
/** Events closer together than this are a continuous (touch) stream; events
 * at/above it are discrete (wheel-like). */
const DISCRETE_MS = 40
/** Right after a flip the very swipe that flipped is often still accelerating,
 * which would pass the rising-velocity test below — ignore that window. */
const REFRACTORY_MS = 200
/** Velocity growth factor that counts as a fresh touch during the tail. */
const RISE = 1.5

type Flip = -1 | 0 | 1

export function createWheelPager(): { feed: (deltaY: number, now: number) => Flip } {
  /** accum: pre-flip, summing toward the first flip of a trackpad gesture
   *  notch: flipped on a discrete notch — more discrete notches keep flipping
   *  swallow: post-flip, eating the same gesture's remaining stream */
  let mode: 'accum' | 'notch' | 'swallow' = 'accum'
  let acc = 0
  let lastTime = -Infinity
  let lastSign = 0
  /** rapid (< DISCRETE_MS) events seen since the last reset/notch flip — the
   * gesture-density signal; 0..1 = sparse enough to still be a wheel */
  let rapidSeen = 0
  // rolling last-two velocities: comparing against their max keeps a single
  // stretched inter-event gap (timing jitter) from reading as a fresh touch
  let vel1 = 0
  let vel2 = 0
  let flipTime = -Infinity

  const feed = (deltaY: number, now: number): Flip => {
    if (deltaY === 0) return 0
    const sign = deltaY > 0 ? 1 : -1
    const dt = now - lastTime
    const vel = Math.abs(deltaY) / Math.max(dt, 1)
    if (dt > GAP_MS || sign !== lastSign) {
      mode = 'accum'
      acc = 0
      rapidSeen = 0
    }
    lastTime = now
    lastSign = sign

    if (
      mode === 'swallow' &&
      now - flipTime > REFRACTORY_MS &&
      Math.max(vel1, vel2) > 0 &&
      vel >= Math.max(vel1, vel2) * RISE
    ) {
      mode = 'accum'
      acc = 0
      rapidSeen = 0
    }
    vel2 = vel1
    vel1 = vel
    if (mode === 'swallow') return 0

    const rapid = dt < DISCRETE_MS
    if (!rapid && Math.abs(deltaY) >= NOTCH && rapidSeen <= 1) {
      // discrete notch in a sparse gesture: flip per notch, and the notch
      // re-asserts wheelness (forgives the tolerated jitter event)
      mode = 'notch'
      acc = 0
      rapidSeen = 0
      flipTime = now
      return sign
    }
    if (mode === 'notch') {
      if (rapid && ++rapidSeen >= 2) {
        // a continuous stream trailing a "notch" means it was a hard flick's
        // opening frame, not a wheel — swallow the rest of the gesture
        mode = 'swallow'
      }
      // isolated jitter or a small slow remnant: eat it, keep paging per notch
      return 0
    }
    if (rapid) rapidSeen++
    acc += deltaY
    if (Math.abs(acc) >= NOTCH) {
      mode = 'swallow'
      acc = 0
      flipTime = now
      return sign
    }
    return 0
  }

  return { feed }
}
