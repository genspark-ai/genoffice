import { describe, expect, it } from 'vitest'
import { OVERLAY_LIMITS, OVERLAY_TONES, validateOverlayShapes } from '../src/slot-overlay'

function check(shapes: unknown): { errors: { field: string; message: string }[]; count: number } {
  const errors: { field: string; message: string }[] = []
  const count = validateOverlayShapes(shapes, 'shapes', errors)
  return { errors, count }
}

describe('overlay shapes', () => {
  it('accepts a frame, a guide, a dot and a label', () => {
    const { errors, count } = check([
      { kind: 'frame', x: 0.1, y: 0.2, width: 0.5, height: 0.2, tone: 'accent', label: 'Title' },
      { kind: 'line', x1: 0, y1: 0.5, x2: 1, y2: 0.5, dashed: true },
      { kind: 'dot', x: 0.9, y: 0.1, tone: 'success', label: 'Check this' },
      { kind: 'label', x: 0.05, y: 0.95, value: 'Draft', tone: 'muted' },
    ])
    expect(errors).toEqual([])
    expect(count).toBe(4)
  })

  it('refuses a coordinate outside the layer, naming the field', () => {
    const { errors } = check([{ kind: 'dot', x: 1.4, y: 0.5 }])
    expect(errors[0]!.field).toBe('shapes[0].x')
    expect(errors[0]!.message).toContain('between 0 and 1')
  })

  it('refuses an invented tone, an unknown shape and an empty layer', () => {
    expect(check([{ kind: 'dot', x: 0.1, y: 0.1, tone: 'neon' }]).errors[0]!.field).toBe(
      'shapes[0].tone',
    )
    expect(check([{ kind: 'circle', x: 0.1, y: 0.1 }]).errors[0]!.field).toBe('shapes[0].kind')
    expect(check([]).errors[0]!.message).toContain('at least one shape')
    expect(check('shapes').errors[0]!.message).toContain('at least one shape')
  })

  it('caps the shapes and the label length', () => {
    const many = Array.from({ length: OVERLAY_LIMITS.maxShapes + 1 }, () => ({
      kind: 'dot',
      x: 0.5,
      y: 0.5,
    }))
    expect(check(many).errors.some((error) => error.message.includes('more than'))).toBe(true)
    const long = check([
      { kind: 'label', x: 0.1, y: 0.1, value: 'x'.repeat(OVERLAY_LIMITS.maxLabelChars + 1) },
    ])
    expect(long.errors[0]!.field).toBe('shapes[0].value')
  })

  it('lists the tones the host paints', () => {
    expect(OVERLAY_TONES).toContain('accent')
    expect(OVERLAY_TONES).toContain('danger')
  })
})
