import { describe, expect, it } from 'vitest'

function isValidRole(role: unknown): boolean {
  return role === 'user' || role === 'assistant'
}

function isValidText(text: unknown): boolean {
  return typeof text === 'string' && text.length <= 200_000
}

describe('slides chat IPC validation', () => {
  it('accepts user/assistant roles', () => {
    expect(isValidRole('user')).toBe(true)
    expect(isValidRole('assistant')).toBe(true)
  })

  it('rejects invalid roles', () => {
    expect(isValidRole('system')).toBe(false)
    expect(isValidRole('')).toBe(false)
    expect(isValidRole(null)).toBe(false)
  })

  it('rejects non-string or oversized text', () => {
    expect(isValidText(123)).toBe(false)
    expect(isValidText('x'.repeat(200_001))).toBe(false)
    expect(isValidText('deck')).toBe(true)
  })
})
