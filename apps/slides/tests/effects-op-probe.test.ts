import { describe, it, expect } from 'vitest'
import { addElement, createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/main/ops'

describe('setEffects op end-to-end probe', () => {
  it('applies an outer shadow through the op registry', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
    })
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setEffects',
          target: { slide: 0, el: el.id },
          effects: { shadow: { color: '#00000066', blurRad: 50800, dist: 38100, dirDeg: 45 } },
        },
      ],
    })
    expect(r.applied).toBe(true)
    expect((el as { shadow?: object }).shadow).toBeTruthy()
    expect(el.anchor.originalXml).toMatch(/outerShdw/)
  })
})
