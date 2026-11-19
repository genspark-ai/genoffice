import { describe, expect, it } from 'vitest'

import { encodeAutoSaveOverride, resolveAutoSave } from '../src/auto-save-pref'

const unset = { on: false, updatedAt: 0 }
const globalOn = { on: true, updatedAt: 1700000000000 }
const globalOff = { on: false, updatedAt: 1700000000000 }

describe('resolveAutoSave', () => {
  it('falls back to the global default when nothing is stored', () => {
    expect(resolveAutoSave(unset, null)).toBe(false)
    expect(resolveAutoSave(globalOn, null)).toBe(true)
    expect(resolveAutoSave(globalOff, '')).toBe(false)
  })

  it('honors legacy 1/0 values only while the global setting was never set', () => {
    expect(resolveAutoSave(unset, '1')).toBe(true)
    expect(resolveAutoSave(unset, '0')).toBe(false)
    expect(resolveAutoSave(globalOn, '0')).toBe(true)
    expect(resolveAutoSave(globalOff, '1')).toBe(false)
  })

  it('honors an override whose base matches the current global stamp', () => {
    expect(resolveAutoSave(globalOn, encodeAutoSaveOverride(false, globalOn))).toBe(false)
    expect(resolveAutoSave(globalOff, encodeAutoSaveOverride(true, globalOff))).toBe(true)
  })

  it('discards an override once the global setting was flipped again', () => {
    const stale = encodeAutoSaveOverride(false, globalOn)
    const flipped = { on: true, updatedAt: globalOn.updatedAt + 1 }
    expect(resolveAutoSave(flipped, stale)).toBe(true)
  })

  it('ignores corrupt or foreign values', () => {
    expect(resolveAutoSave(globalOn, '{')).toBe(true)
    expect(resolveAutoSave(globalOn, '{"on":"yes","base":1700000000000}')).toBe(true)
    expect(resolveAutoSave(globalOff, '[true]')).toBe(false)
  })
})
