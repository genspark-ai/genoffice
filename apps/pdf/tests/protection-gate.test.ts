import { describe, expect, it } from 'vitest'
import { mayChangeProtection } from '../src/renderer/protection-gate'

/**
 * The Protect dialog is the only way to change or lift protection, and its
 * button is gated by this rule: the owner password is what authorizes a change.
 * A user-password session that may modify the content still must not be able to
 * strip the owner's password.
 */
describe('protection changes need the owner password', () => {
  it('lets a plain document gain protection', () => {
    expect(mayChangeProtection(false, null)).toBe(true)
  })

  it('lets the owner change or lift protection', () => {
    expect(mayChangeProtection(false, { unlockedAsOwner: true })).toBe(true)
  })

  it('refuses a user-password session, even one allowed to modify', () => {
    // diskLock.modifying is true here: the /P allows edits, not a password change
    expect(mayChangeProtection(false, { unlockedAsOwner: false })).toBe(false)
  })

  it('refuses every protection change in a read-only session', () => {
    expect(mayChangeProtection(true, { unlockedAsOwner: true })).toBe(false)
    expect(mayChangeProtection(true, null)).toBe(false)
  })
})
