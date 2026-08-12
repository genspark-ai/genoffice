import { describe, expect, it } from 'vitest'
import { signatureDrawingForField } from '../src/renderer/signature-field'
import type { FormWidget } from '../src/renderer/form-catalog'

const target = {
  id: 'sig',
  fieldName: 'signature',
  kind: 'signature',
  pageIndex: 2,
  rect: [100, 200, 300, 260],
  value: '',
  checked: false,
  multiLine: false,
  readOnly: false,
  required: false,
  maxLen: null,
  password: false,
  comb: false,
  textAlignment: 'left',
  buttonValue: '',
  options: [],
  signed: false,
} satisfies FormWidget

describe('signatureDrawingForField', () => {
  it('centers an image signature inside the signature field', () => {
    const drawing = signatureDrawingForField(
      { kind: 'image', image: 'png', width: 400, height: 100 },
      target,
      [0, 0, 0],
    )

    expect(drawing.kind).toBe('image')
    if (drawing.kind !== 'image') return
    expect(drawing.pageIndex).toBe(2)
    expect(drawing.formFieldName).toBe('signature')
    expect(drawing.rect[0]).toBeGreaterThan(100)
    expect(drawing.rect[1]).toBeGreaterThan(200)
    expect(drawing.rect[2]).toBeLessThan(300)
    expect(drawing.rect[3]).toBeLessThan(260)
  })

  it('maps pen strokes into PDF coordinates with the vertical axis flipped', () => {
    const drawing = signatureDrawingForField(
      { kind: 'strokes', width: 100, height: 50, paths: [[0, 0, 100, 50]] },
      target,
      [0.1, 0.2, 0.3],
    )

    expect(drawing.kind).toBe('ink')
    if (drawing.kind !== 'ink') return
    const [startX, startY, endX, endY] = drawing.paths[0]!
    expect(startX).toBeGreaterThanOrEqual(100)
    expect(endX).toBeLessThanOrEqual(300)
    expect(startY).toBeGreaterThan(endY)
    expect(drawing.color).toEqual([0.1, 0.2, 0.3])
    expect(drawing.formFieldName).toBe('signature')
  })
})
