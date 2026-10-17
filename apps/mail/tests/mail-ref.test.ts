import { describe, expect, it } from 'vitest'
import { formatMailRef, isMailRef, parseMailRef } from '../src/shared/mail-ref'

describe('mail-ref', () => {
  it('formats and parses a round trip', () => {
    const ref = formatMailRef('acct-1', 'uid-42')
    expect(ref).toBe('mail://acct-1/uid-42')
    expect(parseMailRef(ref)).toEqual({ accountId: 'acct-1', uid: 'uid-42' })
  })

  it('encodes segments with reserved characters', () => {
    const ref = formatMailRef('acct/1', 'u id')
    expect(parseMailRef(ref)).toEqual({ accountId: 'acct/1', uid: 'u id' })
  })

  it('rejects non-mail refs', () => {
    expect(parseMailRef('https://example.com')).toBeNull()
    expect(parseMailRef('mail://only-account')).toBeNull()
    expect(parseMailRef('')).toBeNull()
    expect(parseMailRef(null)).toBeNull()
    expect(parseMailRef(42)).toBeNull()
    expect(isMailRef('mail://a/b')).toBe(true)
    expect(isMailRef('file.md')).toBe(false)
  })
})
