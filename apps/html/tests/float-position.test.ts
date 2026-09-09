import { describe, expect, it } from 'vitest'
import { floatPosition, parseDeclarations } from '../src/renderer/document/float-position'

const layout = { zoom: 100, offsetX: 0, offsetY: 0, stageWidth: 1000, barWidth: 300, barHeight: 32 }

describe('floatPosition', () => {
  it('sits above the element in stage coordinates', () => {
    expect(floatPosition({ x: 100, y: 200, width: 50, height: 20 }, layout)).toEqual({
      left: 100,
      top: 162,
      below: false,
    })
  })

  it('flips below when there is no room above', () => {
    expect(floatPosition({ x: 10, y: 10, width: 50, height: 20 }, layout)).toEqual({
      left: 10,
      top: 36,
      below: true,
    })
  })

  it('scales with zoom and applies the host offset', () => {
    const pos = floatPosition(
      { x: 100, y: 200, width: 50, height: 20 },
      { ...layout, zoom: 50, offsetX: 40, offsetY: 8 },
    )
    expect(pos).toEqual({ left: 90, top: 70, below: false })
  })

  it('keeps the bar inside the stage', () => {
    expect(floatPosition({ x: 950, y: 200, width: 50, height: 20 }, layout).left).toBe(696)
    expect(floatPosition({ x: -30, y: 200, width: 50, height: 20 }, layout).left).toBe(4)
  })
})

describe('parseDeclarations', () => {
  it('splits declarations and drops junk', () => {
    expect(parseDeclarations('letter-spacing: 2px; box-shadow: 0 1px 2px #000;; nope; :x')).toEqual(
      {
        'letter-spacing': '2px',
        'box-shadow': '0 1px 2px #000',
      },
    )
  })
})
