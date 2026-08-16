/**
 * Passthrough preview rendering: SmartArt previewShapes → group node,
 * OLE previewPicture → picture node filling the frame, media images carry a media flag.
 */
import { describe, it, expect } from 'vitest'
import type { PassthroughElement, PictureElement, Slide, SlideDeck } from '@genoffice/pptx-engine'
import { buildRenderSlide } from '../src/index'
import type { GroupRenderNode, PictureRenderNode } from '../src/render-tree'

const size: SlideDeck['size'] = { cx: 9525 * 1280, cy: 9525 * 720 }

const anchor = { spIndex: 0, originalXml: '<p:graphicFrame/>', range: [0, 0] as [number, number] }
const transform = (x: number, y: number, cx: number, cy: number) => ({
  offset: { x, y, cx, cy },
  rot: 0,
  flipH: false,
  flipV: false,
})

const slideOf = (elements: Slide['elements']): Slide => ({
  path: 'ppt/slides/slide1.xml',
  originalXml: '',
  bodyPrefix: '',
  bodySuffix: '',
  elements,
})

describe('passthrough preview rendering', () => {
  it('smartart previewShapes → group node (child shapes positioned in diagram canvas space)', () => {
    const el: PassthroughElement = {
      id: 'gf1',
      type: 'passthrough',
      kind: 'smartart',
      anchor,
      transform: transform(9525 * 100, 9525 * 50, 9525 * 400, 9525 * 300),
      previewShapes: [
        {
          id: 'd1',
          type: 'shape',
          anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
          transform: transform(0, 0, 9525 * 200, 9525 * 150),
          presetGeometry: 'roundRect',
          fill: { type: 'solid', color: '#4472C4' },
        },
      ],
    }
    const rs = buildRenderSlide(slideOf([el]), size, { fitWidthPx: 1280 })
    expect(rs.nodes.length).toBe(1)
    const g = rs.nodes[0] as GroupRenderNode
    expect(g.type).toBe('group')
    expect(g.sourceId).toBe('gf1')
    expect(g.box.x).toBeCloseTo(100, 0)
    expect(g.children.length).toBe(1)
    expect(g.children[0]!.type).toBe('shape')
    expect(g.children[0]!.box.w).toBeCloseTo(200, 0)
  })

  it('ole previewPicture → picture node fills the frame (ignores the pic own xfrm)', () => {
    const preview: PictureElement = {
      id: 'p1',
      type: 'picture',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: transform(0, 0, 1, 1), // the embedded pic's xfrm should be ignored
      mediaRef: 'ppt/media/image9.png',
    }
    const el: PassthroughElement = {
      id: 'gf2',
      type: 'passthrough',
      kind: 'ole',
      anchor,
      transform: transform(0, 0, 9525 * 300, 9525 * 200),
      previewPicture: preview,
    }
    const rs = buildRenderSlide(slideOf([el]), size, {
      fitWidthPx: 1280,
      media: (ref) => (ref.includes('image9') ? 'data:img' : undefined),
    })
    const n = rs.nodes[0] as PictureRenderNode
    expect(n.type).toBe('picture')
    expect(n.sourceId).toBe('gf2')
    expect(n.dataUrl).toBe('data:img')
    expect(n.box.w).toBeCloseTo(300, 0)
    expect(n.box.h).toBeCloseTo(200, 0)
  })

  it('falls back to placeholder chip when no preview exists', () => {
    const el: PassthroughElement = {
      id: 'gf3',
      type: 'passthrough',
      kind: 'smartart',
      anchor,
      transform: transform(0, 0, 9525 * 100, 9525 * 100),
    }
    const rs = buildRenderSlide(slideOf([el]), size, { fitWidthPx: 1280 })
    expect(rs.nodes[0]!.type).toBe('placeholder-chip')
  })

  it('audio/video pictures carry the media flag', () => {
    const el: PictureElement = {
      id: 'v1',
      type: 'picture',
      anchor: { spIndex: 0, originalXml: '<p:pic/>', range: [0, 0] },
      transform: transform(0, 0, 9525 * 100, 9525 * 100),
      mediaRef: 'ppt/media/poster.png',
      media: { kind: 'video', target: 'ppt/media/media1.mp4' },
    }
    const rs = buildRenderSlide(slideOf([el]), size, { fitWidthPx: 1280 })
    const n = rs.nodes[0] as PictureRenderNode
    expect(n.type).toBe('picture')
    expect(n.media).toBe('video')
  })
})

describe('stale drawing scale-to-frame threshold', () => {
  const mkPreview = (childCy: number): PassthroughElement =>
    ({
      id: 'pt1',
      type: 'passthrough',
      kind: 'smartart',
      anchor,
      transform: transform(0, 0, 1000000, 1000000),
      previewShapes: [
        {
          id: 'sp1',
          type: 'shape',
          anchor,
          transform: transform(0, 0, 1000000, childCy),
          fill: { type: 'solid', color: '#4F81BD' },
        } as any,
      ],
    }) as any
  const groupScaleY = (el: PassthroughElement): number => {
    const rs = buildRenderSlide(slideOf([el]), size, { fitWidthPx: 1280 })
    const g = rs.nodes[0] as GroupRenderNode
    return g.childScaleY ?? 1
  }

  it('a small overshoot (intentional bleed) is left alone', () => {
    expect(groupScaleY(mkPreview(1100000))).toBeCloseTo(1, 3)
  })

  it('a gross overflow (stale cache) compresses into the frame', () => {
    expect(groupScaleY(mkPreview(2500000))).toBeLessThan(0.5)
  })
})
