import { describe, it, expect, beforeAll } from 'vitest'
import {
  addElement,
  addMedia,
  createBlankPptx,
  openPptx,
  getSlideAnimations,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

let opened: OpenedPptx
let videoId: string
let textId: string

beforeAll(async () => {
  opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  addElement(slide, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
    paragraphs: [{ runs: [{ text: 'Hello' }] }],
  })
  videoId = addMedia(opened, 0, {
    kind: 'video',
    bytes: new Uint8Array([0, 0, 0, 0x18, 0x66, 0x74, 0x79, 0x70]),
    ext: 'mp4',
    offset: { x: 0, y: 914400, cx: 1828800, cy: 914400 },
  })!.elementId
  // addMedia reparses the slide, so element ids are reassigned
  textId = opened.deck.slides[0]!.elements.find((e) => e.type !== 'picture')!.id
})

describe('media animations through the ops layer', () => {
  it('addAnimation play on a video writes the media call and the media timeline node', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'addAnimation', target: { slide: 0, el: videoId }, effect: 'play' }],
    })
    expect(r.applied).toBe(true)
    const slide = opened.deck.slides[0]!
    const anims = getSlideAnimations(slide)
    expect(anims).toHaveLength(1)
    expect(anims[0]).toMatchObject({ effect: 'mediaPlay', trigger: 'onClick', durationMs: 0 })
    expect(slide.bodySuffix).toContain('cmd="playFrom(0.0)"')
    expect(slide.bodySuffix).toContain('<p:video>')
  })

  it('rejects media effects on a non-media element', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'addAnimation', target: { slide: 0, el: textId }, effect: 'mediaStop' }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toMatch(/only applies to a video or audio element/)
  })

  it('rejects a mismatched kind cross-check', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'addAnimation',
          target: { slide: 0, el: videoId },
          effect: 'pause',
          kind: 'entrance',
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toMatch(/is an media effect, not entrance|media effect/)
  })
})
