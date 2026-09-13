import { describe, expect, it } from 'vitest'

import { TAB_STRIP_HEIGHT, tabStripOverlay } from '../src/main/title-bar-overlay'

describe('tabStripOverlay', () => {
  it('matches the tab strip band in both themes and its 40px height', () => {
    expect(tabStripOverlay(false)).toEqual({
      color: '#ebebeb',
      symbolColor: '#454746',
      height: TAB_STRIP_HEIGHT,
    })
    expect(tabStripOverlay(true)).toEqual({
      color: '#2a2a2a',
      symbolColor: '#e4e4e4',
      height: TAB_STRIP_HEIGHT,
    })
    expect(TAB_STRIP_HEIGHT).toBe(40)
  })
})
