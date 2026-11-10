import { describe, it, expect } from 'vitest'
import { resolveStroke } from '../src/fill'
import { makeViewport } from '../src/coords'
import type { Stroke } from '@genoffice/pptx-engine'

const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000)

describe('unsupported stroke fills', () => {
  it('drops an image-filled line rather than painting it black', () => {
    const img: Stroke = { fill: { type: 'image', mediaRef: 'ppt/media/a.png' }, width: 12700 }
    expect(resolveStroke(img, vp)).toBeUndefined()
  })

  it('approximates a pattern-filled line with its foreground colour', () => {
    const pat: Stroke = {
      fill: { type: 'pattern', fg: '#336699', bg: '#ffffff', preset: 'dot' },
      width: 12700,
    }
    const resolved = resolveStroke(pat, vp)
    expect(resolved).toBeDefined()
    // the bug: the #000000 default used to win, so every patterned outline rendered black
    expect(resolved!.color).toBe('#336699')
    expect(resolved!.widthPx).toBeGreaterThan(0)
  })
})
