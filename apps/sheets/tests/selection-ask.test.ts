import { describe, expect, it } from 'vitest'
import { isSelectionDrag, selectionAskPosition } from '../src/renderer/ai/selection-ask'

const viewport = { left: 100, top: 50, right: 900, bottom: 650 }

describe('selection Ask AI gesture', () => {
  it('ignores click jitter and accepts a deliberate drag', () => {
    expect(isSelectionDrag({ x: 20, y: 20 }, { x: 22, y: 21 })).toBe(false)
    expect(isSelectionDrag({ x: 20, y: 20 }, { x: 24, y: 20 })).toBe(true)
  })

  it('places the trigger after the drag endpoint when there is room', () => {
    expect(selectionAskPosition({ x: 300, y: 200 }, viewport, 92, 32)).toEqual({
      left: 308,
      top: 208,
    })
  })

  it('flips and clamps the trigger at the grid edges', () => {
    expect(selectionAskPosition({ x: 895, y: 645 }, viewport, 92, 32)).toEqual({
      left: 795,
      top: 605,
    })
  })

  it('uses the measured localized-label width when clamping', () => {
    const position = selectionAskPosition({ x: 895, y: 200 }, viewport, 220, 32)
    expect(position.left).toBe(667)
    expect(position.left + 220).toBeLessThanOrEqual(viewport.right - 4)
  })
})
