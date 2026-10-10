import { describe, expect, it } from 'vitest'

import { tabStripOverlay } from '../src/main/title-bar-overlay'
import { TAB_STRIP_HEIGHT, tabStripHeight } from '../src/shared/tab-drag-geometry'

describe('tabStripOverlay', () => {
  it('matches the tab strip band in both themes and its unscaled height', () => {
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

  it('grows the caption band with the interface scale', () => {
    // Windows and Linux draw the caption buttons over the tab strip, so the
    // overlay band has to be the strip's *rendered* height. The strip is
    // `zoom`ed by `--ui-scale` (#1913); a fixed 40 would leave the buttons
    // floating over the document once the interface is enlarged.
    expect(tabStripOverlay(false, 1.5).height).toBe(60)
    expect(tabStripOverlay(true, 1.5).height).toBe(tabStripHeight(1.5))
    expect(tabStripOverlay(false, 1.5).height).toBe(tabStripOverlay(true, 1.5).height)
  })

  it('defaults to unscaled so an old call site keeps its band', () => {
    // The scale argument is optional: a caller that does not know about #1913
    // still gets the 40px band it had before.
    expect(tabStripOverlay(false).height).toBe(tabStripOverlay(false, 1).height)
  })
})