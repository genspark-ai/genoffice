/**
 * Layout parking — route-B prototype for the large-document lag (#2781).
 *
 * The renderer keeps the whole document as one continuous ProseMirror DOM,
 * so Chromium's style recalc, layout and paint all scale with the document.
 * Parking gives offscreen top-level blocks `content-visibility: auto` with a
 * remembered intrinsic size: Chromium then skips their subtrees entirely
 * between measurement passes, and the caret entering a parked block reveals
 * it (spec: content skipped until "relevant to the user" — on-screen,
 * focused, selected).
 *
 * The measurement pass must read real boxes, so parking is suppressed for
 * the pass (same lifecycle as `measuring-natural`) and re-applied at its
 * end; `contain-intrinsic-size: auto <px>` keeps the last real rendered
 * height so the scrollbar and the page-gap positions hold still while a
 * block is parked.
 *
 * Prototype scope: continuous single-column canvas, zoom 1. Columned
 * canvases translate blocks per column and vertical-text blocks re-wrap;
 * both are excluded by the caller for now.
 */

export const PARKED_CLASS = 'cv-parked'
export const PARK_SUPPRESSED_CLASS = 'parking-suppressed'

/** a screen-coordinate band the scroller shows, grown by headroom */
export interface ParkBand {
  top: number
  bottom: number
}

/** the boxes a pass measured: screen-space top and height */
export interface ParkBox {
  top: number
  height: number
}

/**
 * Which boxes park: everything outside the band. A box straddling the band
 * edge stays unparked — the seam must be real layout, and Chromium reveals
 * skipped content lazily anyway, so a half-visible parked block would only
 * add a reveal right at the scroll edge.
 */
export function parkDecisions(boxes: ParkBox[], band: ParkBand): boolean[] {
  return boxes.map((b) => b.top + b.height <= band.top || b.top >= band.bottom)
}

/**
 * The band to park against: the scroller's viewport grown by one viewport
 * above and below, so normal scrolling lands inside already-real layout and
 * only fast scrolling meets a reveal.
 */
export function parkBand(scroller: HTMLElement): ParkBand {
  const rect = scroller.getBoundingClientRect()
  const headroom = rect.height
  return { top: rect.top - headroom, bottom: rect.bottom + headroom }
}

/** park or reveal each element per its decision, stamping the remembered height */
export function applyParking(
  els: Array<HTMLElement | null | undefined>,
  boxes: ParkBox[],
  decisions: boolean[],
): void {
  const limit = Math.min(els.length, boxes.length, decisions.length)
  for (let i = 0; i < limit; i++) {
    const el = els[i]
    if (!el) continue
    if (decisions[i]) {
      if (!el.classList.contains(PARKED_CLASS)) {
        el.style.setProperty('--park-h', `${Math.round(boxes[i]!.height)}px`)
        el.classList.add(PARKED_CLASS)
      }
    } else if (el.classList.contains(PARKED_CLASS)) {
      el.classList.remove(PARKED_CLASS)
      el.style.removeProperty('--park-h')
    }
  }
}

/** the measurement pass (and fast passes) read real boxes: park nothing */
export function suppressParking(pm: HTMLElement): void {
  pm.classList.add(PARK_SUPPRESSED_CLASS)
}

/** the pass is done sampling: parked blocks may skip again */
export function resumeParking(pm: HTMLElement): void {
  pm.classList.remove(PARK_SUPPRESSED_CLASS)
}

/** whether parking is suppressed right now (a pass is sampling) */
export function parkingSuppressed(pm: HTMLElement): boolean {
  return pm.classList.contains(PARK_SUPPRESSED_CLASS)
}
