import { describe, it, expect } from 'vitest'
import {
  isDuplicateDragModifier,
  isToggleModifier,
  nextSelection,
  type ModifierKeys,
} from '../src/renderer/platform-modifiers'

const keys = (over: Partial<ModifierKeys> = {}): ModifierKeys => ({
  shiftKey: false,
  ctrlKey: false,
  metaKey: false,
  altKey: false,
  ...over,
})

describe('isToggleModifier', () => {
  it('Windows/Linux: Ctrl or Shift toggles, Cmd (Win key) does not', () => {
    expect(isToggleModifier(keys({ ctrlKey: true }), false)).toBe(true)
    expect(isToggleModifier(keys({ shiftKey: true }), false)).toBe(true)
    expect(isToggleModifier(keys({ metaKey: true }), false)).toBe(false)
    expect(isToggleModifier(keys({ altKey: true }), false)).toBe(false)
    expect(isToggleModifier(keys(), false)).toBe(false)
  })

  it('mac: Cmd or Shift toggles; Ctrl stays the system right-click', () => {
    expect(isToggleModifier(keys({ metaKey: true }), true)).toBe(true)
    expect(isToggleModifier(keys({ shiftKey: true }), true)).toBe(true)
    expect(isToggleModifier(keys({ ctrlKey: true }), true)).toBe(false)
    expect(isToggleModifier(keys({ altKey: true }), true)).toBe(false)
  })
})

describe('isDuplicateDragModifier', () => {
  it('Windows/Linux: Ctrl+drag duplicates, Alt+drag does not', () => {
    expect(isDuplicateDragModifier(keys({ ctrlKey: true }), false)).toBe(true)
    expect(isDuplicateDragModifier(keys({ altKey: true }), false)).toBe(false)
    expect(isDuplicateDragModifier(keys(), false)).toBe(false)
  })

  it('mac: Option+drag duplicates, Ctrl+drag does not', () => {
    expect(isDuplicateDragModifier(keys({ altKey: true }), true)).toBe(true)
    expect(isDuplicateDragModifier(keys({ ctrlKey: true }), true)).toBe(false)
    expect(isDuplicateDragModifier(keys({ metaKey: true }), true)).toBe(false)
  })

  it('reads real KeyboardEvent/MouseEvent instances', () => {
    expect(isDuplicateDragModifier(new KeyboardEvent('keydown', { ctrlKey: true }), false)).toBe(
      true,
    )
    expect(isDuplicateDragModifier(new KeyboardEvent('keyup', { ctrlKey: false }), false)).toBe(
      false,
    )
    expect(isToggleModifier(new MouseEvent('click', { metaKey: true }), true)).toBe(true)
  })
})

describe('nextSelection', () => {
  it('plain click replaces the selection', () => {
    expect(nextSelection(['a', 'b'], 'c', false)).toEqual(['c'])
    expect(nextSelection(['a'], 'a', false)).toEqual(['a'])
  })

  it('modifier click adds a new element and removes an already selected one', () => {
    expect(nextSelection(['a'], 'b', true)).toEqual(['a', 'b'])
    expect(nextSelection(['a', 'b'], 'a', true)).toEqual(['b'])
  })

  it('null clears regardless of the modifier', () => {
    expect(nextSelection(['a'], null, true)).toEqual([])
    expect(nextSelection(['a'], null, false)).toEqual([])
  })
})
