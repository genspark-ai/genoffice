import { describe, it, expect } from 'vitest'
import { resolveStroke } from '../src/fill'
import { makeViewport } from '../src/coords'
import type { Stroke } from '@genoffice/pptx-engine'

const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000)

describe('hairline stroke', () => {
  it('keeps explicit width 0 as hairline', () => {
    const s: Stroke = { fill: { type: 'solid', color: '#111' }, width: 0 }
    const r = resolveStroke(s, vp)
    expect(r?.widthPt).toBe(0)
  })
})
