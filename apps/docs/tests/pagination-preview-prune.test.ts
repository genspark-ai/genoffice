import { describe, expect, it } from 'vitest'
import {
  canvasWrapWidth,
  cloneChildGeom,
  prunedCloneHtml,
  shouldPruneClones,
  type CloneChild,
} from '../src/renderer/components/PaginationPreview'

const kid = (partial: Partial<CloneChild> & Pick<CloneChild, 'html' | 'vTop' | 'vBottom'>) => ({
  mt: 0,
  mb: 0,
  zero: false,
  ...partial,
})

/** spacer heights in document order, parsed from the pruned HTML */
const spacerHeights = (html: string): number[] =>
  [...html.matchAll(/pv-prune-spacer" style="[^"]*height:([\d.]+)px/g)].map((m) => Number(m[1]))

// window pad is 2000px (CLONE_PRUNE_PAD): from=10000/to=11000 keeps [8000, 13000]
const FROM = 10000
const TO = 11000

describe('prunedCloneHtml', () => {
  it('keeps in-window blocks verbatim without spacers', () => {
    const kids = [
      kid({ html: '<p id="a"></p>', vTop: 8100, vBottom: 8200 }),
      kid({ html: '<p id="b"></p>', vTop: 8300, vBottom: 8400 }),
    ]
    const out = prunedCloneHtml(kids, FROM, TO)
    expect(out).toBe('<p id="a"></p><p id="b"></p>')
  })

  it('replaces a leading pruned run with a spacer that restores the next block top', () => {
    const kids = [
      kid({ html: '<p id="p"></p>', vTop: 0, vBottom: 500, mb: 5 }),
      kid({ html: '<p id="a"></p>', vTop: 9000, vBottom: 9200, mt: 10 }),
    ]
    const out = prunedCloneHtml(kids, FROM, TO)
    expect(out).not.toContain('id="p"')
    // spacer height re-derives A's border top from flow start: vTop - mt - 0
    expect(spacerHeights(out)).toEqual([8990])
  })

  it('spacer after a mid-run kept block excludes the space before it', () => {
    const kids = [
      kid({ html: '<p id="a"></p>', vTop: 8100, vBottom: 8200, mb: 10 }),
      kid({ html: '<p id="p1"></p>', vTop: 13500, vBottom: 14000 }),
      kid({ html: '<p id="b"></p>', vTop: 8500, vBottom: 8600, mt: 20 }),
    ]
    // contrived order aside: any kept block after a pruned run gets a spacer
    // based on the previous kept block's margin edge (8200 + 10)
    const out = prunedCloneHtml(kids, FROM, TO)
    expect(spacerHeights(out)).toEqual([8500 - 20 - (8200 + 10)])
  })

  it('advances the flow base across zero-height markers (regression: stale spacer base)', () => {
    const kids = [
      kid({ html: '<p id="p1"></p>', vTop: 1000, vBottom: 2000 }),
      kid({ html: '<span id="z"></span>', vTop: 3000, vBottom: 3000, zero: true }),
      kid({ html: '<p id="p2"></p>', vTop: 4000, vBottom: 5000 }),
      kid({ html: '<p id="b"></p>', vTop: 9000, vBottom: 9500, mt: 20 }),
    ]
    const out = prunedCloneHtml(kids, FROM, TO)
    // marker is always kept even though it sits outside the window
    expect(out).toContain('id="z"')
    expect(out).not.toContain('id="p1"')
    expect(out).not.toContain('id="p2"')
    // spacer1 places the marker at 3000; spacer2 must start from the marker's
    // flow position (3000), not from the pre-marker base — the buggy stale
    // base would re-add the 3000px already consumed and shift B a page down
    expect(spacerHeights(out)).toEqual([3000, 9000 - 20 - 3000])
  })
})

describe('shouldPruneClones', () => {
  it('keeps full clones for a document that fits both budgets', () => {
    expect(shouldPruneClones(20, 500, 20_000)).toBe(false)
  })

  it('prunes once pages x blocks passes the block budget', () => {
    expect(shouldPruneClones(300, 500, 1_000)).toBe(true)
  })

  it('prunes a page-heavy document whose few blocks carry a huge DOM', () => {
    // under the block budget (112 x 869), 9.7 M elements over the element one
    expect(shouldPruneClones(112, 869, 87_039)).toBe(true)
  })

  it('prunes a 1.7 M-element preview that sits under both former budgets', () => {
    expect(shouldPruneClones(107, 859, 15_919)).toBe(true)
  })
})

describe('canvasWrapWidth', () => {
  // A4 with 1134-twip margins: 642.5333px nominal, Blink lays the canvas out at 642.5390625
  const NOMINAL = 642.5333333333333

  it('returns the snapped width of a full-width canvas block', () => {
    expect(canvasWrapWidth([22.5, 642.5390625, 701, 642.5390625], NOMINAL)).toBe(642.5390625)
  })

  it('ignores indented blocks, over-wide tables and empty canvases', () => {
    expect(canvasWrapWidth([600.2, 701, 0], NOMINAL)).toBeNull()
    expect(canvasWrapWidth([], NOMINAL)).toBeNull()
  })

  it('prefers the block closest to the nominal width', () => {
    expect(canvasWrapWidth([642.5, 642.546875, 642.5390625], NOMINAL)).toBe(642.5390625)
  })
})

describe('cloneChildGeom', () => {
  it('positions a block-backed child on the pass geometry, a block-less one on its rect', () => {
    const raw = { vTop: 100, h: 40 }
    expect(cloneChildGeom(undefined, raw)).toEqual(raw)
    // the lead fold moved the first block's space before out of its top
    expect(cloneChildGeom({ top: 0, height: 120, domHeight: 40, leadFoldPx: 100 }, raw)).toEqual({
      vTop: 100,
      h: 40,
    })
    // a page-relative V float: the applied shift was taken out of the slicer top,
    // the clone paints the shifted box
    expect(
      cloneChildGeom({ top: 300, height: 50, domHeight: 50, shiftAppliedPx: 80 }, raw),
    ).toEqual({ vTop: 380, h: 50 })
    // a joined paragraph tail extended the head's DOM box; footnote extras never do
    expect(
      cloneChildGeom({ top: 100, height: 160, domHeight: 60, footnoteExtraPx: 100 }, raw),
    ).toEqual({ vTop: 100, h: 60 })
  })
})
