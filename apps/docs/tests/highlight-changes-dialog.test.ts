/**
 * Word's Highlight Changes dialog: "Highlight changes on screen" maps onto the
 * markup view, and only switches it when the tick disagrees with what is shown.
 */
import { describe, expect, it } from 'vitest'
import {
  displayForHighlight,
  displayForPrint,
  highlightsOnScreen,
} from '../src/renderer/components/HighlightChangesDialog'

describe('highlight changes on screen', () => {
  it('is ticked in the views that draw markup', () => {
    expect(highlightsOnScreen('all')).toBe(true)
    expect(highlightsOnScreen('simple')).toBe(true)
    expect(highlightsOnScreen('none')).toBe(false)
    expect(highlightsOnScreen('original')).toBe(false)
  })

  it('leaves the view alone when the tick already matches it', () => {
    expect(displayForHighlight('simple', true)).toBeNull()
    expect(displayForHighlight('original', false)).toBeNull()
  })

  it('switches to All Markup when ticked and No Markup when cleared', () => {
    expect(displayForHighlight('none', true)).toBe('all')
    expect(displayForHighlight('original', true)).toBe('all')
    expect(displayForHighlight('all', false)).toBe('none')
    expect(displayForHighlight('simple', false)).toBe('none')
  })

  it('a print follows the screen until the printed-document box is ticked or cleared', () => {
    for (const mode of ['all', 'simple', 'none', 'original'] as const) {
      expect(displayForPrint(mode, null)).toBe(mode)
      expect(displayForPrint(mode, true)).toBe('all')
    }
    expect(displayForPrint('all', false)).toBe('none')
    // these already print the clean sheet
    expect(displayForPrint('simple', false)).toBe('simple')
    expect(displayForPrint('original', false)).toBe('original')
  })
})
