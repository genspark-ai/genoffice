import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  DEFAULT_SPIN_COUNT,
  MAX_SPIN_COUNT,
  hashProtectionPassword,
  resolveSpinCount,
  verifyProtectionPassword,
} from '../src/protection'

afterEach(() => {
  vi.restoreAllMocks()
})

describe('resolveSpinCount', () => {
  it('returns the default for missing values', () => {
    expect(resolveSpinCount(undefined)).toBe(DEFAULT_SPIN_COUNT)
    expect(resolveSpinCount(null)).toBe(DEFAULT_SPIN_COUNT)
    expect(resolveSpinCount('')).toBe(DEFAULT_SPIN_COUNT)
    expect(DEFAULT_SPIN_COUNT).toBe(100000)
  })

  it('accepts valid custom counts', () => {
    expect(resolveSpinCount(0)).toBe(0)
    expect(resolveSpinCount(1000)).toBe(1000)
    expect(resolveSpinCount('50000')).toBe(50000)
    expect(resolveSpinCount(MAX_SPIN_COUNT)).toBe(MAX_SPIN_COUNT)
  })

  it('clamps huge counts to the maximum', () => {
    expect(MAX_SPIN_COUNT).toBe(1000000)
    expect(resolveSpinCount(10_000_000)).toBe(MAX_SPIN_COUNT)
    expect(resolveSpinCount(99_999_999)).toBe(MAX_SPIN_COUNT)
  })

  it('rejects negative, non-integer, and non-numeric values', () => {
    expect(() => resolveSpinCount(-1)).toThrow(/spinCount/)
    expect(() => resolveSpinCount(-100000)).toThrow(/spinCount/)
    expect(() => resolveSpinCount(1.5)).toThrow(/spinCount/)
    expect(() => resolveSpinCount(NaN)).toThrow(/spinCount/)
    expect(() => resolveSpinCount(Infinity)).toThrow(/spinCount/)
    expect(() => resolveSpinCount('abc')).toThrow(/spinCount/)
    expect(() => resolveSpinCount({})).toThrow(/spinCount/)
  })
})

describe('verifyProtectionPassword DoS guard', () => {
  const salt = 'AAAAAAAAAAAAAAAAAAAAAA=='

  it('fails closed on absurd spinCount without hashing or throwing', async () => {
    const digest = vi.spyOn(globalThis.crypto.subtle, 'digest')
    await expect(
      verifyProtectionPassword('pw', {
        hash: 'eA==',
        salt,
        spinCount: 99_999_999,
        algorithmSid: 14,
      }),
    ).resolves.toBe(false)
    expect(digest).not.toHaveBeenCalled()
  })

  it('rejects negative and non-integer spinCount without calling Subtle.digest', async () => {
    const digest = vi.spyOn(globalThis.crypto.subtle, 'digest')
    await expect(
      verifyProtectionPassword('pw', { hash: 'eA==', salt, spinCount: -5, algorithmSid: 14 }),
    ).rejects.toThrow(/spinCount/)
    await expect(
      verifyProtectionPassword('pw', { hash: 'eA==', salt, spinCount: 1.5, algorithmSid: 14 }),
    ).rejects.toThrow(/spinCount/)
    await expect(
      verifyProtectionPassword('pw', { hash: 'eA==', salt, spinCount: NaN, algorithmSid: 14 }),
    ).rejects.toThrow(/spinCount/)
    expect(digest).not.toHaveBeenCalled()
  })

  it('returns false on missing salt instead of hashing or throwing', async () => {
    const digest = vi.spyOn(globalThis.crypto.subtle, 'digest')
    await expect(
      verifyProtectionPassword('pw', { hash: 'eA==', spinCount: 1000, algorithmSid: 14 }),
    ).resolves.toBe(false)
    expect(digest).not.toHaveBeenCalled()
  })

  it('still verifies a password hashed with a small spinCount', async () => {
    const creds = await hashProtectionPassword('correct', 1000)
    expect(creds.spinCount).toBe(1000)
    await expect(verifyProtectionPassword('correct', { ...creds, algorithmSid: 14 })).resolves.toBe(
      true,
    )
    await expect(verifyProtectionPassword('wrong', { ...creds, algorithmSid: 14 })).resolves.toBe(
      false,
    )
  })
})
