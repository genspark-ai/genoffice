import { describe, expect, it } from 'vitest'
import { drawingOutline, outlineHit, toSavedShape } from '../src/renderer/shape-hit'
import type { DrawingInput } from '../src/shared/ipc'

const base = { pageIndex: 0, color: [0, 0, 0] as [number, number, number], width: 2 }
const hit = (d: DrawingInput, x: number, y: number, tol = 4) =>
  outlineHit(drawingOutline(d)!, x, y, tol)

describe('drawingOutline + outlineHit', () => {
  it('hits an ink stroke along its path, not inside its bounding box', () => {
    const ink: DrawingInput = { ...base, kind: 'ink', paths: [[0, 0, 100, 0, 100, 100]] }
    expect(hit(ink, 50, 3)).toBe(true)
    expect(hit(ink, 98, 60)).toBe(true)
    expect(hit(ink, 40, 60)).toBe(false)
  })

  it('hits a rectangle and an ellipse on the border only', () => {
    const rect: DrawingInput = { ...base, kind: 'rect', rect: [0, 0, 200, 100] }
    expect(hit(rect, 100, 1)).toBe(true)
    expect(hit(rect, 100, 50)).toBe(false)
    const ellipse: DrawingInput = { ...base, kind: 'ellipse', rect: [0, 0, 200, 100] }
    expect(hit(ellipse, 200, 50)).toBe(true)
    expect(hit(ellipse, 100, 50)).toBe(false)
    expect(hit(ellipse, 2, 2)).toBe(false)
  })

  it('reaches to the edge of a thick stroke, pending or saved', () => {
    const ink: DrawingInput = { ...base, width: 8, kind: 'ink', paths: [[0, 0, 100, 0]] }
    expect(hit(ink, 50, 4.4, 0.5)).toBe(true)
    expect(hit(ink, 50, 5, 0.5)).toBe(false)
    const saved = toSavedShape(
      {
        id: '3R',
        annotationType: 15,
        rect: [0, -10, 100, 10],
        inkLists: [[0, 0, 100, 0]],
        borderStyle: { width: 8 },
      },
      0,
    )!
    expect(outlineHit(saved.outline, 50, 4.4, 0.5)).toBe(true)
    expect(outlineHit(saved.outline, 50, 5, 0.5)).toBe(false)
  })

  it('hits the closing edge of a polygon', () => {
    const polygon: DrawingInput = {
      ...base,
      kind: 'polygon',
      vertices: [
        [0, 0],
        [100, 0],
        [100, 100],
      ],
    }
    expect(hit(polygon, 50, 50)).toBe(true)
    expect(hit(polygon, 80, 30)).toBe(false)
  })

  it('leaves notes, images and links to their own controls', () => {
    expect(
      drawingOutline({ kind: 'link', pageIndex: 0, rect: [0, 0, 10, 10], url: 'https://x.test' }),
    ).toBeNull()
  })
})

describe('toSavedShape', () => {
  it('reads an Ink annotation with its stroke geometry', () => {
    const shape = toSavedShape(
      {
        id: '12R',
        annotationType: 15,
        rect: [0, 0, 100, 100],
        inkLists: [Float32Array.from([0, 0, 100, 100])],
      },
      3,
    )!
    expect(shape).toMatchObject({ pageIndex: 3, objNum: 12, type: 'ink' })
    expect(outlineHit(shape.outline, 50, 50, 3)).toBe(true)
    expect(outlineHit(shape.outline, 90, 10, 3)).toBe(false)
  })

  it('hits a saved rectangle and ellipse on the stroke, which sits inside the padded rect', () => {
    // a 2pt border drawn at [100, 100, 300, 200]: the save path pads /Rect by width + 2
    const data = { id: '9R', rect: [96, 96, 304, 204], borderStyle: { width: 2 } }
    const square = toSavedShape({ ...data, annotationType: 5 }, 0)!
    expect(outlineHit(square.outline, 200, 100, 0.5)).toBe(true)
    expect(outlineHit(square.outline, 200, 96, 0.5)).toBe(true)
    expect(outlineHit(square.outline, 200, 108, 0.5)).toBe(false)
    expect(outlineHit(square.outline, 200, 150, 0.5)).toBe(false)
    const circle = toSavedShape({ ...data, annotationType: 6 }, 0)!
    expect(outlineHit(circle.outline, 300, 150, 0.5)).toBe(true)
    expect(outlineHit(circle.outline, 200, 150, 0.5)).toBe(false)
    expect(outlineHit(circle.outline, 100, 100, 0.5)).toBe(false)
  })

  it('falls back to the rect when a shape carries no geometry', () => {
    const shape = toSavedShape({ id: '7R', annotationType: 15, rect: [10, 10, 50, 50] }, 0)!
    expect(outlineHit(shape.outline, 30, 30, 1)).toBe(true)
  })

  it('skips markups, widgets and annotations without an object reference', () => {
    expect(toSavedShape({ id: '7R', annotationType: 9, rect: [0, 0, 1, 1] }, 0)).toBeNull()
    expect(toSavedShape({ id: 'annot_1', annotationType: 15, rect: [0, 0, 1, 1] }, 0)).toBeNull()
  })
})
