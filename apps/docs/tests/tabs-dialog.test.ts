/** Word's Tabs dialog: the stops a paragraph is given and what is written back for them. */
import { describe, expect, it } from 'vitest'
import type { TabStop } from '@genoffice/docx-engine'
import { parseTabStops, tabStopsAttr, withTabStop } from '../src/renderer/components/tab-stops'

describe('tab stops of a paragraph', () => {
  it('sets a stop in order and replaces the one already at that position', () => {
    const stops: TabStop[] = [{ pos: 1440, val: 'left' }]
    expect(withTabStop(stops, { pos: 720, val: 'center' }).map((s) => s.pos)).toEqual([720, 1440])
    expect(withTabStop(stops, { pos: 1440, val: 'right', leader: 'dot' })).toEqual([
      { pos: 1440, val: 'right', leader: 'dot' },
    ])
  })

  it('reads the attribute apart and survives one that is not a list', () => {
    const raw = JSON.stringify([
      { pos: 1440, val: 'left' },
      { pos: 50, val: 'center', rel: 'margin' },
    ])
    expect(parseTabStops(raw).stops).toEqual([{ pos: 1440, val: 'left' }])
    expect(parseTabStops(raw).relStops).toHaveLength(1)
    expect(parseTabStops('{').stops).toEqual([])
    expect(parseTabStops(null).stops).toEqual([])
  })

  it('writes the edited stops, clears a removed style stop and keeps the position tabs', () => {
    const raw = JSON.stringify([
      { pos: 720, val: 'left', inherited: true },
      { pos: 1440, val: 'left' },
      { pos: 50, val: 'center', rel: 'margin' },
    ])
    expect(JSON.parse(tabStopsAttr(raw, [{ pos: 2880, val: 'right' }]) ?? '')).toEqual([
      { pos: 720, val: 'clear' },
      { pos: 2880, val: 'right' },
      { pos: 50, val: 'center', rel: 'margin' },
    ])
    expect(tabStopsAttr(JSON.stringify([{ pos: 1440, val: 'left' }]), [])).toBeNull()
  })

  it('keeps a clear from an earlier removal, unless a stop is set at its position', () => {
    const raw = JSON.stringify([
      { pos: 720, val: 'clear' },
      { pos: 1440, val: 'left' },
    ])
    expect(JSON.parse(tabStopsAttr(raw, [{ pos: 2880, val: 'right' }]) ?? '')).toEqual([
      { pos: 720, val: 'clear' },
      { pos: 2880, val: 'right' },
    ])
    // Clear All leaves the clear standing
    expect(JSON.parse(tabStopsAttr(raw, []) ?? '')).toEqual([{ pos: 720, val: 'clear' }])
    expect(JSON.parse(tabStopsAttr(raw, [{ pos: 720, val: 'center' }]) ?? '')).toEqual([
      { pos: 720, val: 'center' },
    ])
  })
})
