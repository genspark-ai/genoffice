import { describe, expect, it } from 'vitest'
import { newPasteCascade, pageKey, pasteShiftPx, recordPaste } from '../src/main/paste-cascade'

const src = pageKey(1, 0)
const other = pageKey(1, 1)

describe('element paste cascade', () => {
  it('first paste onto another slide lands at the source position', () => {
    const clip = newPasteCascade(src)
    expect(pasteShiftPx(clip, other)).toBe(0)
  })

  it('pasting back onto the copy slide shifts past the original', () => {
    const clip = newPasteCascade(src)
    expect(pasteShiftPx(clip, src)).toBe(16)
    recordPaste(clip, src)
    expect(pasteShiftPx(clip, src)).toBe(32)
  })

  it('repeat pastes onto another slide cascade from the exact position', () => {
    const clip = newPasteCascade(src)
    recordPaste(clip, other)
    expect(pasteShiftPx(clip, other)).toBe(16)
    recordPaste(clip, other)
    expect(pasteShiftPx(clip, other)).toBe(32)
  })

  it('each new slide gets its own first in-place landing', () => {
    const clip = newPasteCascade(src)
    recordPaste(clip, other)
    expect(pasteShiftPx(clip, pageKey(1, 2))).toBe(0)
  })

  it('cut leaves no original: pasting back onto the source slide lands in place', () => {
    const clip = newPasteCascade(null)
    expect(pasteShiftPx(clip, src)).toBe(0)
    recordPaste(clip, src)
    expect(pasteShiftPx(clip, src)).toBe(16)
  })

  it('the same slide index in another window is a different page', () => {
    const clip = newPasteCascade(src)
    expect(pasteShiftPx(clip, pageKey(2, 0))).toBe(0)
  })
})
