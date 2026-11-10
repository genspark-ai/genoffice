import { describe, it, expect } from 'vitest'
import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

// The Slides app inserts a blank slide above slide 1 as one transaction; the
// plan validates every op against the pre-transaction deck, so the move must
// target the pre-existing slide, even when it is the only one.
describe('blank slide above the first slide in one transaction', () => {
  it('lands the new slide at index 0 of a single-slide deck', async () => {
    const opened = await openPptx(await createBlankPptx())
    const original = opened.deck.slides[0]!.path
    const r = runTxn(opened, {
      ops: [
        { op: 'addBlankSlide', target: { slide: 0 } },
        { op: 'moveSlide', target: { slide: 0 }, to: 1 },
      ],
    })
    expect(r.applied).toBe(true)
    expect(opened.deck.slides.map((s) => s.path)).toHaveLength(2)
    expect(opened.deck.slides[1]!.path).toBe(original)
    expect(opened.deck.slides[0]!.path).not.toBe(original)
  })
})
