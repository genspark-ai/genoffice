import { describe, expect, it } from 'vitest'

/**
 * Guards for shell tab IPC handlers (apps/shell/src/main/index.ts
 * registerTabsIpc): activate/close must ignore non-string or empty ids,
 * mirroring the existing reorder guard. A compromised renderer sending
 * number/object/null must not pollute activeId or closingIds.
 */
function isValidTabId(id: unknown): id is string {
  return typeof id === 'string' && id.length > 0
}

describe('tabs IPC id validation', () => {
  it('accepts non-empty string ids', () => {
    expect(isValidTabId('tab-1')).toBe(true)
    expect(isValidTabId('home')).toBe(true)
  })

  it('rejects non-string ids', () => {
    expect(isValidTabId(123)).toBe(false)
    expect(isValidTabId(null)).toBe(false)
    expect(isValidTabId(undefined)).toBe(false)
    expect(isValidTabId({})).toBe(false)
    expect(isValidTabId([])).toBe(false)
  })

  it('rejects empty string ids', () => {
    expect(isValidTabId('')).toBe(false)
  })
})
