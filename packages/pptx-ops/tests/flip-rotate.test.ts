import { describe, it, expect } from 'vitest'
import { addElement, createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

describe('flipElements on rotated shape', () => {
  it('keeps offset unchanged, only toggles flip', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'rect',
      offset: { x: 100000, y: 100000, cx: 1000000, cy: 500000 },
    })
    el.transform.rot = 45 * 60000
    const before = { ...el.transform.offset }
    const r = runTxn(opened, {
      ops: [{ op: 'flipElements', target: { slide: 0 }, els: [el.id], axis: 'h' }],
    })
    expect(r.applied).toBe(true)
    const after = slide.elements.find((x) => x.id === el.id)!.transform
    expect(after.flipH).toBe(true)
    expect(after.offset).toEqual(before)
  })
})
