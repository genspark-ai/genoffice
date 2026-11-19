/**
 * The pagination pass measures anchored drawings (text boxes, pictures, groups)
 * as it finds them; a picture still decoding or a group laid out late is a
 * zero-height box the pass skips, and the cuts it published are stale the
 * moment the box gets its size. The watch records the boxes' heights at pass
 * time and reports the ones that changed afterwards, so the canvas schedules
 * another pass instead of the preview printing the stale one.
 */

/** anchored drawing boxes the pass measures as floats (pagination-measure) */
export const FLOAT_BOX_SELECTOR =
  '.doc-protected-floating .doc-textbox, .doc-img-float .doc-img-wrap, .doc-inline-img-anchor > img'

/** layout px, independent of the canvas zoom (CSS zoom leaves offsetHeight unscaled) */
const heightOfEl = (el: HTMLElement): number => el.offsetHeight

export function floatBoxHeights(
  pm: HTMLElement,
  heightOf: (el: HTMLElement) => number = heightOfEl,
): Map<HTMLElement, number> {
  const out = new Map<HTMLElement, number>()
  for (const el of pm.querySelectorAll<HTMLElement>(FLOAT_BOX_SELECTOR))
    if (!el.closest('.page-gap, .page-float-host, .page-repeat-header')) out.set(el, heightOf(el))
  return out
}

/** boxes whose height moved from the recorded one by more than `tol` px */
export function changedFloatBoxes(
  recorded: ReadonlyMap<HTMLElement, number>,
  heightOf: (el: HTMLElement) => number = heightOfEl,
  tol = 0.5,
): HTMLElement[] {
  const out: HTMLElement[] = []
  for (const [el, h] of recorded) {
    if (!el.isConnected) continue
    if (Math.abs(heightOf(el) - h) > tol) out.push(el)
  }
  return out
}

export function zeroHeightCount(heights: ReadonlyMap<HTMLElement, number>): number {
  let n = 0
  for (const h of heights.values()) if (h <= 0) n++
  return n
}

export interface FloatWatchRecord {
  boxes: number
  zeroHeight: number
}

export class FloatWatch {
  private readonly ro: ResizeObserver | null
  private recorded = new Map<HTMLElement, number>()
  private readonly loading = new Set<HTMLImageElement>()
  private readonly onLoad = () => this.check()

  constructor(
    private readonly onChange: (changed: HTMLElement[]) => void,
    private readonly heightOf: (el: HTMLElement) => number = heightOfEl,
  ) {
    this.ro = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(() => this.check())
  }

  /** snapshot the boxes the pass just measured and watch them */
  record(pm: HTMLElement): FloatWatchRecord {
    this.unwatch()
    this.recorded = floatBoxHeights(pm, this.heightOf)
    for (const el of this.recorded.keys()) {
      this.ro?.observe(el)
      if (el instanceof HTMLImageElement && !el.complete) {
        el.addEventListener('load', this.onLoad, { once: true })
        this.loading.add(el)
      }
    }
    return { boxes: this.recorded.size, zeroHeight: zeroHeightCount(this.recorded) }
  }

  /** compare now; changed boxes are re-recorded so one change reports once */
  check(): HTMLElement[] {
    const changed = changedFloatBoxes(this.recorded, this.heightOf)
    if (changed.length === 0) return changed
    for (const el of changed) this.recorded.set(el, this.heightOf(el))
    this.onChange(changed)
    return changed
  }

  private unwatch(): void {
    this.ro?.disconnect()
    for (const img of this.loading) img.removeEventListener('load', this.onLoad)
    this.loading.clear()
  }

  dispose(): void {
    this.unwatch()
    this.recorded = new Map()
  }
}
