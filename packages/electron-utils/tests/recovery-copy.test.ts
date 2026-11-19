import { describe, expect, it } from 'vitest'

import { recoveryCopyDiskNewer } from '../src/recovery-copy'

describe('recoveryCopyDiskNewer', () => {
  it('is false when the recovery copy postdates the original (plain crash)', () => {
    expect(recoveryCopyDiskNewer(2_000, 1_000)).toBe(false)
  })

  it('is true when the original was modified after the copy was written', () => {
    expect(recoveryCopyDiskNewer(1_000, 2_000)).toBe(true)
  })

  it('treats equal timestamps as disk-newer so the user still sees the warning', () => {
    expect(recoveryCopyDiskNewer(1_000, 1_000)).toBe(true)
  })
})
