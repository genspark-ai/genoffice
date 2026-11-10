import { describe, it, expect } from 'vitest'
import { resolveStroke } from '../src/fill'
import { makeViewport } from '../src/coords'
import type { Stroke } from '@genoffice/pptx-engine'

const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000)
const dashOf = (dash: string) => {
  const s: Stroke = { fill: { type: 'solid', color: '#111' }, width: 9525, dash }
  return resolveStroke(s, vp)?.dash
}

describe('shape dash presets', () => {
  it('distinguishes sys variants from base presets', () => {
    expect(dashOf('dash')).toEqual([4, 3])
    expect(dashOf('sysDash')).toEqual([3, 1])
    expect(dashOf('dot')).toEqual([1, 3])
    expect(dashOf('sysDot')).toEqual([1, 1])
    expect(dashOf('sysDashDot')).toEqual([3, 1, 1, 1])
    expect(dashOf('sysDashDotDot')).toEqual([3, 1, 1, 1, 1, 1])
  })
})
