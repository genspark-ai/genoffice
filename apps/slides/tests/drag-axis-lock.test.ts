import { describe, it, expect } from 'vitest'
import { constrainAxis } from '../src/renderer/drag-axis-lock'

describe('constrainAxis (Shift+drag)', () => {
  it('keeps the horizontal move when |dx| dominates', () => {
    expect(constrainAxis(40, 12)).toEqual({ dx: 40, dy: 0 })
    expect(constrainAxis(-40, 12)).toEqual({ dx: -40, dy: 0 })
  })

  it('keeps the vertical move when |dy| dominates', () => {
    expect(constrainAxis(5, -30)).toEqual({ dx: 0, dy: -30 })
    expect(constrainAxis(-5, 30)).toEqual({ dx: 0, dy: 30 })
  })

  it('ties (including no movement) resolve to horizontal', () => {
    expect(constrainAxis(20, 20)).toEqual({ dx: 20, dy: 0 })
    expect(constrainAxis(-20, 20)).toEqual({ dx: -20, dy: 0 })
    expect(constrainAxis(0, 0)).toEqual({ dx: 0, dy: 0 })
  })

  it('re-evaluates as the deltas change so the user can switch axis mid-drag', () => {
    const path: Array<[number, number]> = [
      [10, 2],
      [30, 8],
      [30, 25],
      [30, 45],
      [70, 45],
    ]
    const axes = path.map(([dx, dy]) => (constrainAxis(dx, dy).dy === 0 ? 'x' : 'y'))
    expect(axes).toEqual(['x', 'x', 'x', 'y', 'x'])
  })

  it('never alters the surviving component', () => {
    expect(constrainAxis(3.25, -0.5)).toEqual({ dx: 3.25, dy: 0 })
    expect(constrainAxis(0.5, -3.25)).toEqual({ dx: 0, dy: -3.25 })
  })
})
