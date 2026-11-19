/** addModel3d / setModel3dView: one transaction each, never offered to the AI. */
import { describe, it, expect } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  solidPng,
  type PictureElement,
} from '@genoffice/pptx-engine'
import { OP_DOCS, opVocabulary } from '../src/op-docs'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

/** A .glb holding only the JSON chunk `{}`. */
const EMPTY_GLB = Buffer.from('Z2xURgIAAAAYAAAABAAAAEpTT057fSAg', 'base64').toString('base64')
const OFFSET = { offset: { x: 0, y: 0, cx: 2000000, cy: 2000000 } }

describe('3D model ops', () => {
  it('inserts a model, then turns it and stores its rendering', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = runTxn(opened, {
      ops: [{ op: 'addModel3d', target: { slide: 0 }, bytes: EMPTY_GLB, ext: 'glb', ...OFFSET }],
    })
    expect(added.failures ?? []).toEqual([])
    const model = opened.deck.slides[0]!.elements.at(-1) as PictureElement
    expect(model.model3d?.rot).toEqual({ ax: 0, ay: 0, az: 0 })

    const turned = runTxn(opened, {
      ops: [
        {
          op: 'setModel3dView',
          target: { slide: 0, el: model.id },
          ax: 600000,
          ay: -1200000,
          az: 0,
          poster: Buffer.from(solidPng(2, 2, [1, 2, 3])).toString('base64'),
          posterExt: 'png',
        },
      ],
    })
    expect(turned.failures ?? []).toEqual([])
    const after = opened.deck.slides[0]!.elements.at(-1) as PictureElement
    expect(after.model3d?.rot).toEqual({ ax: 600000, ay: -1200000, az: 0 })
  })

  it('refuses to turn an element that is not a model', async () => {
    const opened = await openPptx(await createBlankPptx())
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 100, cy: 100 },
    })!
    const r = runTxn(opened, {
      ops: [{ op: 'setModel3dView', target: { slide: 0, el: el.id }, ax: 0, ay: 0, az: 0 }],
    })
    expect(r.failures?.[0]?.error).toContain('is not a 3D model')
  })

  it('is not offered to the AI', () => {
    for (const name of ['addModel3d', 'setModel3dView']) {
      expect(OP_DOCS[name]!.aiCallable).toBe(false)
      expect(opVocabulary()).not.toContain(name)
    }
  })
})
