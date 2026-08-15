import { describe, expect, it } from 'vitest'
import { CARD_PIN_ALIGN, layoutMarginCards } from '../src/renderer/note-margin-layout'

const fixedHeight =
  (h: number) =>
  (_key: string): number =>
    h

describe('layoutMarginCards', () => {
  it('aligns a lone card with its pin', () => {
    const tops = layoutMarginCards([{ key: 'a', pinY: 200 }], fixedHeight(80), null)
    expect(tops.get('a')).toBe(200 - CARD_PIN_ALIGN)
  })

  it('clamps cards near the top edge to minTop', () => {
    const tops = layoutMarginCards([{ key: 'a', pinY: 5 }], fixedHeight(80), null)
    expect(tops.get('a')).toBe(4)
  })

  it('pushes overlapping cards down without an anchor', () => {
    const tops = layoutMarginCards(
      [
        { key: 'a', pinY: 100 },
        { key: 'b', pinY: 120 },
      ],
      fixedHeight(80),
      null,
    )
    expect(tops.get('a')).toBe(100 - CARD_PIN_ALIGN)
    // b wants 120-14=106 but a occupies until 86+80=166, plus the 10px gap
    expect(tops.get('b')).toBe(100 - CARD_PIN_ALIGN + 80 + 10)
  })

  it('leaves far-apart cards at their pin-aligned positions', () => {
    const tops = layoutMarginCards(
      [
        { key: 'a', pinY: 100 },
        { key: 'b', pinY: 500 },
      ],
      fixedHeight(80),
      null,
    )
    expect(tops.get('a')).toBe(100 - CARD_PIN_ALIGN)
    expect(tops.get('b')).toBe(500 - CARD_PIN_ALIGN)
  })

  it('keeps the anchored card fixed and pushes the card above it up', () => {
    const tops = layoutMarginCards(
      [
        { key: 'a', pinY: 100 },
        { key: 'b', pinY: 120 },
      ],
      fixedHeight(80),
      'b',
    )
    expect(tops.get('b')).toBe(120 - CARD_PIN_ALIGN)
    // a is pulled up so it ends 10px above b, even past minTop if needed
    expect(tops.get('a')).toBe(120 - CARD_PIN_ALIGN - 10 - 80)
  })

  it('pushes cards below the anchor down', () => {
    const tops = layoutMarginCards(
      [
        { key: 'a', pinY: 100 },
        { key: 'b', pinY: 130 },
        { key: 'c', pinY: 160 },
      ],
      fixedHeight(60),
      'a',
    )
    expect(tops.get('a')).toBe(100 - CARD_PIN_ALIGN)
    expect(tops.get('b')).toBe(100 - CARD_PIN_ALIGN + 60 + 10)
    expect(tops.get('c')).toBe(100 - CARD_PIN_ALIGN + 2 * (60 + 10))
  })

  it('sorts entries by pin position regardless of input order', () => {
    const tops = layoutMarginCards(
      [
        { key: 'low', pinY: 300 },
        { key: 'high', pinY: 40 },
      ],
      fixedHeight(50),
      null,
    )
    expect(tops.get('high')).toBe(40 - CARD_PIN_ALIGN)
    expect(tops.get('low')).toBe(300 - CARD_PIN_ALIGN)
  })

  it('falls back to the plain pass when the anchor key is unknown', () => {
    const tops = layoutMarginCards([{ key: 'a', pinY: 100 }], fixedHeight(80), 'gone')
    expect(tops.get('a')).toBe(100 - CARD_PIN_ALIGN)
  })
})
