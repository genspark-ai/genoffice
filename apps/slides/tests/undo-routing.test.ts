import { describe, expect, it } from 'vitest'
import { isTextUndoTarget, shouldRouteHistoryToDeck } from '../src/renderer/undo-routing'

function target(tagName: string, attributes: Record<string, string> = {}) {
  return {
    tagName,
    isContentEditable: false,
    getAttribute: (name: string) => attributes[name] ?? null,
  }
}

describe('Slides history shortcut routing', () => {
  it('routes a cleared, untouched AI composer to deck undo and redo', () => {
    const textarea = target('TEXTAREA', {
      'data-slides-ai-input': 'true',
      'data-deck-undo-ready': 'true',
    })
    expect(isTextUndoTarget(textarea)).toBe(true)
    expect(shouldRouteHistoryToDeck(textarea)).toBe(true)
  })

  it('preserves native history after the user types in the AI composer', () => {
    const textarea = target('TEXTAREA', {
      'data-slides-ai-input': 'true',
      'data-deck-undo-ready': 'false',
    })
    expect(shouldRouteHistoryToDeck(textarea)).toBe(false)
  })

  it('preserves native history in ordinary inputs and editable text', () => {
    expect(shouldRouteHistoryToDeck(target('INPUT'))).toBe(false)
    expect(isTextUndoTarget({ isContentEditable: true })).toBe(true)
  })
})
