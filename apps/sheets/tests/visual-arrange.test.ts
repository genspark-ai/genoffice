import { describe, expect, it } from 'vitest'

import {
  alignBoxes,
  changedPositions,
  distributeBoxes,
  normalizeDegrees,
  reorderZ,
  rotatedAabb,
} from '../src/renderer/visual-arrange'

const box = (id: string, x: number, y: number, width: number, height: number) => ({
  id,
  x,
  y,
  width,
  height,
})

describe('alignBoxes', () => {
  const boxes = [box('a', 10, 20, 100, 50), box('b', 40, 60, 20, 10), box('c', 0, 100, 60, 40)]

  it('aligns to the selection extremes and the bounding-box center', () => {
    expect(alignBoxes(boxes, 'left')).toEqual([
      { id: 'a', x: 0, y: 20 },
      { id: 'b', x: 0, y: 60 },
    ])
    expect(alignBoxes(boxes, 'right')).toEqual([
      { id: 'b', x: 90, y: 60 },
      { id: 'c', x: 50, y: 100 },
    ])
    expect(alignBoxes(boxes, 'center')).toEqual([
      { id: 'a', x: 5, y: 20 },
      { id: 'b', x: 45, y: 60 },
      { id: 'c', x: 25, y: 100 },
    ])
    expect(alignBoxes(boxes, 'top')).toEqual([
      { id: 'b', x: 40, y: 20 },
      { id: 'c', x: 0, y: 20 },
    ])
    expect(alignBoxes(boxes, 'bottom')).toEqual([
      { id: 'a', x: 10, y: 90 },
      { id: 'b', x: 40, y: 130 },
    ])
    expect(alignBoxes(boxes, 'middle')).toEqual([
      { id: 'a', x: 10, y: 55 },
      { id: 'b', x: 40, y: 75 },
      { id: 'c', x: 0, y: 60 },
    ])
  })

  it('needs at least two objects', () => {
    expect(alignBoxes([boxes[0]!], 'left')).toEqual([])
  })
})

describe('distributeBoxes', () => {
  it('spaces the inner objects evenly between the outer two', () => {
    const boxes = [box('a', 0, 0, 10, 10), box('c', 90, 5, 10, 10), box('b', 20, 1, 20, 10)]
    expect(distributeBoxes(boxes, 'horizontal')).toEqual([{ id: 'b', x: 40, y: 1 }])
    const vertical = [box('a', 0, 0, 10, 10), box('b', 3, 12, 10, 30), box('c', 9, 100, 10, 10)]
    expect(distributeBoxes(vertical, 'vertical')).toEqual([{ id: 'b', x: 3, y: 40 }])
  })

  it('needs at least three objects and reports only moved ones', () => {
    expect(distributeBoxes([box('a', 0, 0, 1, 1), box('b', 5, 0, 1, 1)], 'horizontal')).toEqual([])
    const even = [box('a', 0, 0, 10, 10), box('b', 20, 0, 10, 10), box('c', 40, 0, 10, 10)]
    expect(distributeBoxes(even, 'horizontal')).toEqual([])
  })
})

describe('rotatedAabb', () => {
  it('swaps the sides at 90 degrees and grows them at 45', () => {
    const quarter = rotatedAabb(100, 50, 90)
    expect(quarter.width).toBeCloseTo(50)
    expect(quarter.height).toBeCloseTo(50 + 50)
    const diagonal = rotatedAabb(100, 100, 45)
    expect(diagonal.width).toBeCloseTo(Math.SQRT2 * 100)
    expect(diagonal.height).toBeCloseTo(Math.SQRT2 * 100)
    expect(rotatedAabb(100, 50, -270).width).toBeCloseTo(quarter.width)
  })

  it('normalizes degrees into [0, 360)', () => {
    expect(normalizeDegrees(-90)).toBe(270)
    expect(normalizeDegrees(450)).toBe(90)
  })
})

describe('reorderZ', () => {
  const order = ['a', 'b', 'c', 'd']

  it('moves one step or to either end', () => {
    expect(reorderZ(order, 'b', 'front')).toEqual(['a', 'c', 'd', 'b'])
    expect(reorderZ(order, 'c', 'back')).toEqual(['c', 'a', 'b', 'd'])
    expect(reorderZ(order, 'b', 'forward')).toEqual(['a', 'c', 'b', 'd'])
    expect(reorderZ(order, 'b', 'backward')).toEqual(['b', 'a', 'c', 'd'])
  })

  it('is a no-op at the ends and for unknown ids', () => {
    expect(reorderZ(order, 'd', 'front')).toEqual(order)
    expect(reorderZ(order, 'a', 'backward')).toEqual(order)
    expect(reorderZ(order, 'zz', 'front')).toEqual(order)
  })

  it('reports only the positions that changed', () => {
    const next = reorderZ(order, 'b', 'front')
    expect([...changedPositions(order, next)]).toEqual([
      ['c', 1],
      ['d', 2],
      ['b', 3],
    ])
  })
})
