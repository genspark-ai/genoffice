import { describe, expect, it } from 'vitest'

import { placeFolderMenu } from '../src/renderer/src/folder-menu-position'

const viewport = { width: 1000, height: 700 }
const menu = { width: 180, height: 220 }

describe('placeFolderMenu', () => {
  it('opens to the right of the anchor when it fits', () => {
    expect(placeFolderMenu({ x: 120, y: 200, above: 180 }, menu, viewport)).toEqual({
      left: 120,
      top: 200,
    })
  })

  it('right-aligns a "..." button menu under its button', () => {
    expect(placeFolderMenu({ x: 400, y: 200, above: 180, align: 'end' }, menu, viewport)).toEqual({
      left: 400 - 180,
      top: 200,
    })
  })

  it('clamps a right-aligned menu inside the left edge', () => {
    expect(placeFolderMenu({ x: 50, y: 200, above: 180, align: 'end' }, menu, viewport).left).toBe(
      8,
    )
  })

  it('keeps the menu inside the left edge when the anchor is near it', () => {
    expect(placeFolderMenu({ x: 2, y: 200, above: 180 }, menu, viewport).left).toBe(8)
  })

  it('keeps the menu inside the right edge', () => {
    expect(placeFolderMenu({ x: 990, y: 200, above: 180 }, menu, viewport).left).toBe(
      1000 - 180 - 8,
    )
  })

  it('flips above the anchor when there is no room below', () => {
    expect(placeFolderMenu({ x: 120, y: 600, above: 580 }, menu, viewport).top).toBe(580 - 220)
  })

  it('clamps the top when even the flipped position would be out of view', () => {
    expect(placeFolderMenu({ x: 120, y: 690, above: 5 }, menu, viewport).top).toBe(8)
  })

  it('never goes negative when the menu is wider than the viewport', () => {
    expect(
      placeFolderMenu({ x: 50, y: 10, above: 0 }, { width: 2000, height: 10 }, viewport).left,
    ).toBe(8)
  })
})
