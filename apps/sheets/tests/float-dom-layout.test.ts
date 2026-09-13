import { describe, expect, it } from 'vitest'

import { floatDomLayout } from '../src/renderer/WorkbookVisuals'

describe('floatDomLayout', () => {
  it('grows the rect by the inset Univer shaves so the inner box equals the frame', () => {
    // in-cell icon: twoCellAnchor to = 304800 x 142875 EMU = 32 x 15 px.
    const layout = floatDomLayout({ width: 32, height: 15, marginX: 0, marginY: 0 })
    // Univer renders the inner box as (width - 4) x (height - 4), offset (2, 2).
    expect(layout.width - 4).toBe(32)
    expect(layout.height - 4).toBe(15)
    expect(layout.marginX + 2).toBe(0)
    expect(layout.marginY + 2).toBe(0)
  })

  it('keeps a non-zero in-cell offset anchored at the same pixel', () => {
    const layout = floatDomLayout({ width: 300, height: 180.5, marginX: 12.25, marginY: 7 })
    expect(layout).toEqual({ width: 304, height: 184.5, marginX: 10.25, marginY: 5 })
  })
})
