/**
 * Per-pass counters for the float geometry channels a pagination pass writes
 * and a later pass reads back (`--tblp-dy`, the `doc-table-float-flow` class,
 * `--page-float-dy`, the float-carry spacer). A converged document must stop
 * writing them: a pass that still flips one is a pass that moved real layout
 * after measuring it. Read through `__pageDebug.layoutHealth`; instrumentation
 * only, no layout effect.
 */
export interface LayoutHealthPass {
  tblpDy: number
  floatFlow: number
  pageFloatDy: number
  carry: number
}

const newPass = (): LayoutHealthPass => ({ tblpDy: 0, floatFlow: 0, pageFloatDy: 0, carry: 0 })

const MAX_PASSES = 64

class LayoutHealth {
  /** one entry per pass of the current follow-up chain, oldest first */
  passes: LayoutHealthPass[] = [newPass()]

  /** `fresh`: a pass that was not self-scheduled starts a new chain */
  beginPass(fresh: boolean): void {
    if (fresh) this.passes = []
    if (this.passes.length >= MAX_PASSES) this.passes.shift()
    this.passes.push(newPass())
  }

  count(kind: keyof LayoutHealthPass, n = 1): void {
    if (n > 0) this.passes[this.passes.length - 1][kind] += n
  }

  /** writes in every pass after the chain's first: zero once the first pass converged */
  flipsAfterFirst(): LayoutHealthPass {
    return this.passes.slice(1).reduce((acc, p) => {
      acc.tblpDy += p.tblpDy
      acc.floatFlow += p.floatFlow
      acc.pageFloatDy += p.pageFloatDy
      acc.carry += p.carry
      return acc
    }, newPass())
  }

  snapshot(): { passes: LayoutHealthPass[]; flipsAfterFirst: LayoutHealthPass } {
    return { passes: this.passes.map((p) => ({ ...p })), flipsAfterFirst: this.flipsAfterFirst() }
  }
}

export const layoutHealth = new LayoutHealth()

/** the display float shift already applied to a box (px) */
export const appliedFloatShift = (el: HTMLElement): number =>
  parseFloat(el.dataset.pageFloatDy ?? '0') || 0

/** the one writer of the display float shift (--page-float-dy): idempotent,
 *  a change counts once on the layout-health pass */
// in the editor the float slot sits on the chrome rotator (the layer that turns
// the shape); the shift must ride there or the box slides out of it
export function floatShiftTarget(el: HTMLElement): HTMLElement {
  const parent = el.parentElement
  return parent?.classList.contains('box-rotator') ? parent : el
}

export function writeFloatShift(el: HTMLElement, next: number): void {
  const applied = appliedFloatShift(el)
  const target = floatShiftTarget(el)
  if (Math.abs(next) < 0.5) {
    if (el.dataset.pageFloatDy) {
      layoutHealth.count('pageFloatDy')
      target.style.removeProperty('--page-float-dy')
      delete el.dataset.pageFloatDy
    }
  } else if (Math.abs(next - applied) > 0.5 || !el.dataset.pageFloatDy) {
    layoutHealth.count('pageFloatDy')
    target.style.setProperty('--page-float-dy', `${next.toFixed(1)}px`)
    el.dataset.pageFloatDy = String(next)
  }
}
