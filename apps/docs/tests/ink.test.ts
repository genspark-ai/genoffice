import { describe, expect, it } from 'vitest'
import type { Block, InkInfo } from '@genoffice/docx-engine'
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  annotationsFromParsed,
  buildInkImages,
  decodeInkPayload,
  encodeInkPayload,
  strokeBounds,
  strokeLayers,
  strokeHitTest,
  type InkAnnotation,
  type InkRaster,
  type InkStroke,
} from '../src/renderer/editor/ink'

const STROKE: InkStroke = {
  tool: 'pen',
  color: 'C00000',
  width: 2,
  points: [
    { x: 10, y: 5 },
    { x: 60, y: 25 },
  ],
}

describe('ink payload round-trip', () => {
  it('encodes and decodes a stroke', () => {
    const decoded = decodeInkPayload(encodeInkPayload(STROKE))
    expect(decoded).toEqual(STROKE)
  })

  it('rejects malformed / foreign payloads', () => {
    expect(decodeInkPayload('not json')).toBeNull()
    expect(decodeInkPayload('{"v":99,"points":[[0,0]]}')).toBeNull()
    expect(decodeInkPayload('{"v":1,"points":[]}')).toBeNull()
  })
})

describe('annotationsFromParsed', () => {
  const base: Omit<InkInfo, 'payload' | 'dataUrl'> = {
    anchorIndex: 3,
    offsetXPx: 8,
    offsetYPx: -4,
    widthPx: 100,
    heightPx: 40,
  }

  it('restores editable strokes from the payload', () => {
    const [ann] = annotationsFromParsed([
      { ...base, dataUrl: 'data:image/png;base64,x', payload: encodeInkPayload(STROKE) },
    ])
    expect(ann.anchorIndex).toBe(3)
    expect(ann.stroke).toEqual(STROKE)
    expect(ann.image).toBeUndefined()
  })

  it('falls back to an image annotation when the payload is unreadable', () => {
    const [ann] = annotationsFromParsed([
      { ...base, dataUrl: 'data:image/png;base64,x', payload: null },
    ])
    expect(ann.stroke).toBeUndefined()
    expect(ann.image).toEqual({
      dataUrl: 'data:image/png;base64,x',
      offsetXPx: 8,
      offsetYPx: -4,
      widthPx: 100,
      heightPx: 40,
    })
  })
})

describe('stroke geometry', () => {
  it('bounds include the line width padding', () => {
    const b = strokeBounds(STROKE)
    expect(b.x).toBe(10 - 2)
    expect(b.y).toBe(5 - 2)
    expect(b.w).toBe(50 + 4)
    expect(b.h).toBe(20 + 4)
  })

  it('hit test matches near the segment and misses far away', () => {
    expect(strokeHitTest(STROKE, { x: 35, y: 15 }, 3)).toBe(true)
    expect(strokeHitTest(STROKE, { x: 35, y: 40 }, 3)).toBe(false)
  })
})

describe('buildInkImages', () => {
  const fakeRaster = (stroke: InkStroke): InkRaster => {
    const b = strokeBounds(stroke)
    return { base64: 'PNGDATA', widthPx: b.w, heightPx: b.h, offsetXPx: b.x, offsetYPx: b.y }
  }

  it('maps docx anchors to save-block indexes and encodes the payload', () => {
    const anns: InkAnnotation[] = [{ id: 'a', anchorIndex: 7, stroke: STROKE }]
    const out = buildInkImages(anns, new Map([[7, 2]]), fakeRaster)
    expect(out).toHaveLength(1)
    expect(out[0].blockIndex).toBe(2)
    expect(out[0].base64).toBe('PNGDATA')
    expect(decodeInkPayload(out[0].payload!)).toEqual(STROKE)
  })

  it('drops annotations whose anchor paragraph was deleted', () => {
    const anns: InkAnnotation[] = [{ id: 'a', anchorIndex: 7, stroke: STROKE }]
    expect(buildInkImages(anns, new Map(), fakeRaster)).toHaveLength(0)
  })

  it('carries restored image annotations through unchanged', () => {
    const anns: InkAnnotation[] = [
      {
        id: 'a',
        anchorIndex: 1,
        image: {
          dataUrl: 'data:image/png;base64,QUJD',
          offsetXPx: 5,
          offsetYPx: 6,
          widthPx: 70,
          heightPx: 30,
        },
      },
    ]
    const out = buildInkImages(anns, new Map([[1, 0]]), fakeRaster)
    expect(out).toEqual([
      {
        blockIndex: 0,
        base64: 'QUJD',
        widthPx: 70,
        heightPx: 30,
        offsetXPx: 5,
        offsetYPx: 6,
      },
    ])
  })
})

describe('pmDocToSavePlan saveBlockIndexByDocx', () => {
  const para = (docxIndex: number, text: string): Block => ({
    id: `b${docxIndex}`,
    docxIndex,
    originalXml: `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`,
    type: 'paragraph',
    runs: [{ text }],
  })
  const pmPara = (docxIndex: number | null, text: string): PmNode => ({
    type: 'docParagraph',
    attrs: { docxIndex },
    content: [{ type: 'text', text }],
  })

  it('maps unchanged, edited and shifted paragraphs to their save positions', () => {
    const originals = [para(0, 'a'), para(1, 'b')]
    const doc: PmNode = {
      type: 'doc',
      content: [
        pmPara(null, 'inserted'), // new paragraph at the top shifts everything
        pmPara(0, 'a'), // unchanged -> original
        pmPara(1, 'b CHANGED'), // edited -> generated, anchor consumed
      ],
    }
    const plan = pmDocToSavePlan(doc, originals)
    expect(plan.saveBlocks[1]).toEqual({ kind: 'original', docxIndex: 0 })
    expect(plan.saveBlocks[2].kind).toBe('generated')
    expect(plan.saveBlockIndexByDocx.get(0)).toBe(1)
    expect(plan.saveBlockIndexByDocx.get(1)).toBe(2)
  })

  it('deleted paragraphs get no mapping', () => {
    const originals = [para(0, 'a'), para(1, 'b')]
    const doc: PmNode = { type: 'doc', content: [pmPara(0, 'a')] }
    const plan = pmDocToSavePlan(doc, originals)
    expect(plan.saveBlockIndexByDocx.get(0)).toBe(0)
    expect(plan.saveBlockIndexByDocx.has(1)).toBe(false)
    expect(plan.deletedCount).toBe(1)
  })
})

describe('pencil layers', () => {
  const pts = Array.from({ length: 30 }, (_, i) => ({ x: i * 4, y: Math.sin(i / 3) * 10 }))
  const pencil = { tool: 'pencil' as const, color: '000000', width: 8, points: pts }

  it('grain already drawn stays put while the live stroke grows', () => {
    const full = strokeLayers(pencil)
    const partial = strokeLayers({ ...pencil, points: pts.slice(0, 12) })
    // only the tail point moves: its normal is re-estimated once a next point exists
    for (let l = 0; l < 3; l++)
      for (let i = 0; i < 11; i++) expect(partial[l].points[i]).toEqual(full[l].points[i])
  })

  it('bounds cover every grain layer, not just the polyline', () => {
    const b = strokeBounds(pencil)
    for (const layer of strokeLayers(pencil))
      for (const p of layer.points) {
        expect(p.x - layer.width / 2).toBeGreaterThanOrEqual(b.x)
        expect(p.y - layer.width / 2).toBeGreaterThanOrEqual(b.y)
        expect(p.x + layer.width / 2).toBeLessThanOrEqual(b.x + b.w)
        expect(p.y + layer.width / 2).toBeLessThanOrEqual(b.y + b.h)
      }
  })

  it('payload round-trips the pencil kind', () => {
    expect(decodeInkPayload(encodeInkPayload(pencil))?.tool).toBe('pencil')
  })
})
