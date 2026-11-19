/**
 * A pass that ran before an anchored drawing had its size published stale cuts
 * (a letterhead whose logo group was still decoding exported as header/footer
 * only). The watch reports the boxes whose height moved after the pass; the
 * canvas re-runs the pass and the new cuts carry the float.
 */
import { describe, expect, it, vi } from 'vitest'
import type { SectionSettings } from '@genoffice/docx-engine'
import {
  FloatWatch,
  changedFloatBoxes,
  floatBoxHeights,
  zeroHeightCount,
} from '../src/renderer/editor/float-watch'
import { pageCutSignature, type SectionGeom, type SliceOutputs } from '../src/renderer/pagination'
import { slicePass, type PassContext } from '../src/renderer/pagination-pass'

function canvasWith(boxHeights: number[]): { pm: HTMLElement; boxes: HTMLElement[] } {
  const pm = document.createElement('div')
  const boxes = boxHeights.map((h) => {
    const wrap = document.createElement('div')
    wrap.className = 'doc-protected-floating'
    const box = document.createElement('div')
    box.className = 'doc-textbox'
    box.dataset.h = String(h)
    wrap.appendChild(box)
    pm.appendChild(wrap)
    return box
  })
  const gap = document.createElement('div')
  gap.className = 'page-gap'
  gap.innerHTML =
    '<div class="doc-protected-floating"><div class="doc-textbox" data-h="99"></div></div>'
  pm.appendChild(gap)
  // the watch ignores boxes that left the canvas (isConnected)
  document.body.appendChild(pm)
  return { pm, boxes }
}
const heightOf = (el: HTMLElement) => parseFloat(el.dataset.h ?? '0') || 0

describe('float-watch', () => {
  it('records the anchored boxes (not the gap widgets) and counts the zero-height ones', () => {
    const { pm, boxes } = canvasWith([0, 40])
    const rec = floatBoxHeights(pm, heightOf)
    expect([...rec.keys()]).toEqual(boxes)
    expect(zeroHeightCount(rec)).toBe(1)
    expect(changedFloatBoxes(rec, heightOf)).toEqual([])
    boxes[0].dataset.h = '120'
    expect(changedFloatBoxes(rec, heightOf)).toEqual([boxes[0]])
  })

  it('reports a late box once and stays quiet until it moves again', () => {
    const { pm, boxes } = canvasWith([0, 40])
    const onChange = vi.fn()
    const watch = new FloatWatch(onChange, heightOf)
    expect(watch.record(pm)).toEqual({ boxes: 2, zeroHeight: 1 })
    expect(watch.check()).toEqual([])
    boxes[0].dataset.h = '120'
    expect(watch.check()).toEqual([boxes[0]])
    expect(onChange).toHaveBeenCalledTimes(1)
    expect(watch.check()).toEqual([])
    expect(watch.record(pm)).toEqual({ boxes: 2, zeroHeight: 0 })
    watch.dispose()
    boxes[1].dataset.h = '10'
    expect(watch.check()).toEqual([])
  })

  it('a flow whose float gains height after the first pass re-slices to a longer page', () => {
    const NO_HF = { headerPx: 0, footerPx: 0 }
    const geom: SectionGeom = {
      contentHeight: 800,
      contentWidth: 600,
      marginBottomPx: 60,
      forceBreak: false,
    }
    const out = (): SliceOutputs => ({
      rowFills: [],
      rowSplits: [],
      floatVShifts: [],
      floatPullUps: [],
      floatCarries: [],
      floatFlows: [],
      floatSplits: [],
      oversizeClips: [],
    })
    const ctx = (): PassContext => ({
      sections: [],
      section: { pageHeight: 15840, marginBottom: 1440 } as SectionSettings,
      endnoteItems: [],
      endnoteSeparatorPx: 0,
      hfHeightsOf: () => [NO_HF],
      singleHfPx: NO_HF,
      singleGeom: geom,
      evenOddHf: false,
      titlePg: false,
      dropCols: false,
      factor: 1,
      out: out(),
    })
    const blocks = () => [{ top: 0, height: 60, el: document.createElement('p') }]
    const el = document.body.appendChild(document.createElement('div'))
    // pass 1: the logo group has no size yet, measureBlocks reports no float
    const first = slicePass(
      { blocks: blocks(), totalHeight: 60, floats: [], sectBreaks: new Set() },
      ctx(),
    )
    expect(first.cuts.map((c) => c.end)).toEqual([60])
    // the box gets its size: the watch reports it, the pass runs again
    const recorded = new Map([[el, 0]])
    const changed = changedFloatBoxes(recorded, () => 400)
    expect(changed).toEqual([el])
    const second = slicePass(
      {
        blocks: blocks(),
        totalHeight: 60,
        floats: [{ el, top: 10, height: 400, anchorTop: 0, pinned: false, pageRelV: false }],
        sectBreaks: new Set(),
      },
      ctx(),
    )
    expect(second.cuts[0].end).toBeGreaterThan(60)
    expect(second.cuts[0].boxes).toEqual([0])
    expect(pageCutSignature(second.cuts)).not.toBe(pageCutSignature(first.cuts))
  })
})
