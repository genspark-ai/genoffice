import { beforeEach, describe, expect, it } from 'vitest'
import type { RenderNode } from '@genoffice/pptx-render'
import {
  defaultShapeStyle,
  saveDefaultShapeStyle,
  shapeStyleOf,
} from '../src/renderer/default-shape'

const shape = (over: Record<string, unknown>) =>
  ({ type: 'shape', sourceId: 's', fill: { kind: 'none' }, ...over }) as unknown as RenderNode

describe('default shape style', () => {
  beforeEach(() => localStorage.clear())

  it('keeps the alpha byte of translucent fills and lines', () => {
    expect(
      shapeStyleOf(
        shape({
          fill: { kind: 'solid', color: '#11223380' },
          stroke: { color: '44556680', widthPt: 1.5 },
        }),
      ),
    ).toEqual({ fillColor: '#11223380', stroke: { color: '#44556680', widthPt: 1.5 } })
  })

  it('a translucent fill alone still qualifies', () => {
    expect(shapeStyleOf(shape({ fill: { kind: 'solid', color: '#C43E1C40' } }))).toEqual({
      fillColor: '#C43E1C40',
    })
  })

  it('rejects connectors, gradients without a line, and non-hex colors', () => {
    expect(
      shapeStyleOf(shape({ line: { points: [] }, fill: { kind: 'solid', color: '#112233' } })),
    ).toBeUndefined()
    expect(
      shapeStyleOf(shape({ fill: { kind: 'gradient', stops: [], angleDeg: 0 } })),
    ).toBeUndefined()
    expect(shapeStyleOf(shape({ fill: { kind: 'solid', color: 'red' } }))).toBeUndefined()
    expect(
      shapeStyleOf({ type: 'picture', sourceId: 'p' } as unknown as RenderNode),
    ).toBeUndefined()
  })

  it('round-trips through localStorage and falls back to the factory fill', () => {
    expect(defaultShapeStyle()).toEqual({ fillColor: '#C43E1C' })
    saveDefaultShapeStyle({ fillColor: '#11223380' })
    expect(defaultShapeStyle()).toEqual({ fillColor: '#11223380' })
  })
})
