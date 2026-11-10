import { describe, expect, it } from 'vitest'
import { isFromInspector } from '../src/renderer/preview/inspector-protocol'

describe('isFromInspector', () => {
  it('accepts well-formed inspector messages', () => {
    expect(isFromInspector({ type: 'gx:ready', version: 3, title: 't', docHeight: 10 })).toBe(true)
    expect(isFromInspector({ type: 'gx:zoom', version: 1, delta: 5 })).toBe(true)
  })

  it('rejects forged or malformed shapes', () => {
    expect(isFromInspector(null)).toBe(false)
    expect(isFromInspector('gx:ready')).toBe(false)
    // Unknown tag, even with the gx: prefix.
    expect(isFromInspector({ type: 'gx:evil', version: 1 })).toBe(false)
    // Version is required and numeric (the App matches it against live parse maps).
    expect(isFromInspector({ type: 'gx:ready', title: 't' })).toBe(false)
    expect(isFromInspector({ type: 'gx:ready', version: '3' })).toBe(false)
    expect(isFromInspector({ type: 'gx:ready', version: Number.NaN })).toBe(false)
  })
})
