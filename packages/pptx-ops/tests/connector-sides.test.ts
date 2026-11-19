import { describe, it, expect } from 'vitest'
import { addElement, createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

describe('addConnector rotated sides', () => {
  it('picks the visual east side of a 90-degree shape', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const from = addElement(slide, { kind: 'rect', offset: { x: 0, y: 0, cx: 200000, cy: 100000 } })
    const to = addElement(slide, {
      kind: 'rect',
      offset: { x: 1000000, y: 0, cx: 200000, cy: 100000 },
    })
    from.transform.rot = 90 * 60000
    const r = runTxn(opened, {
      ops: [{ op: 'addConnector', target: { slide: 0 }, from: from.id, to: to.id }],
    })
    expect(r.applied).toBe(true)
    const after = (r.records[0] as any).after
    expect(after.fromSide).toBe('top')
    expect(after.toSide).toBe('left')
  })
})
