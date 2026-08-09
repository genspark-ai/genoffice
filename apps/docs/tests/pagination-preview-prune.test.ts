import { describe, expect, it } from 'vitest'
import { prunedCloneHtml, type CloneChild } from '../src/renderer/components/PaginationPreview'

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
