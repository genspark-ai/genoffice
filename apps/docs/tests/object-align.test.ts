import { describe, expect, it } from 'vitest'
import {
  alignedBoxes,
  boundsOf,
  effectiveAlignFrame,
  type Box,
} from '../src/renderer/editor/object-align'

const box = (left: number, top: number, width: number, height: number): Box => ({
  left,
  top,
  width,
  height,
})

const BOXES = [box(10, 100, 40, 20), box(200, 10, 100, 60), box(90, 300, 20, 100)]
const PAGE = box(0, 0, 800, 1000)

describe('aligning objects', () => {
  it('lines the objects up on an edge of what surrounds them and leaves the other axis alone', () => {
    const own = boundsOf(BOXES)
    expect(own).toEqual(box(10, 10, 290, 390))
    expect(alignedBoxes(BOXES, 'left', own).map((b) => [b.left, b.top])).toEqual([
      [10, 100],
      [10, 10],
      [10, 300],
    ])
    expect(alignedBoxes(BOXES, 'right', own).map((b) => b.left + b.width)).toEqual([300, 300, 300])
    expect(alignedBoxes(BOXES, 'center', own).map((b) => b.left + b.width / 2)).toEqual([
      155, 155, 155,
    ])
    expect(alignedBoxes(BOXES, 'top', own).map((b) => [b.left, b.top])).toEqual([
      [10, 10],
      [200, 10],
      [90, 10],
    ])
    expect(alignedBoxes(BOXES, 'middle', own).map((b) => b.top + b.height / 2)).toEqual([
      205, 205, 205,
    ])
    expect(alignedBoxes(BOXES, 'bottom', own).map((b) => b.top + b.height)).toEqual([400, 400, 400])
  })

  it('lines one object up against the page', () => {
    expect(alignedBoxes([BOXES[0]], 'center', PAGE)[0]).toEqual(box(380, 100, 40, 20))
    expect(alignedBoxes([BOXES[0]], 'bottom', PAGE)[0]).toEqual(box(10, 980, 40, 20))
  })

  it('distributes with equal gaps, the outer objects staying where they are', () => {
    const across = alignedBoxes(BOXES, 'distribute-h', boundsOf(BOXES))
    // 290 wide, 160 of objects: two gaps of 65, in left-to-right order
    expect(across.map((b) => b.left)).toEqual([10, 200, 115])
    expect(across.map((b) => b.top)).toEqual([100, 10, 300])
    const down = alignedBoxes(BOXES, 'distribute-v', boundsOf(BOXES))
    // 390 tall, 180 of objects: two gaps of 105, in top-to-bottom order
    expect(down.map((b) => b.top)).toEqual([175, 10, 300])
  })

  it('distributes over the page edge to edge, and centres one object alone', () => {
    expect(alignedBoxes(BOXES, 'distribute-h', PAGE).map((b) => b.left)).toEqual([0, 700, 360])
    expect(alignedBoxes([BOXES[1]], 'distribute-v', PAGE)[0].top).toBe(470)
  })

  it('uses the page for one object even when objects are what to align against', () => {
    expect(effectiveAlignFrame('objects', 1)).toBe('page')
    expect(effectiveAlignFrame('objects', 2)).toBe('objects')
    expect(effectiveAlignFrame('margin', 1)).toBe('margin')
  })
})
