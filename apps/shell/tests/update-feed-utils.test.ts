import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'

// the utils are a dependency-free CommonJS script shared with CI; load via
// createRequire since this test file is ESM
const require = createRequire(import.meta.url)
const {
  ymlVersion,
  semverNewer,
  assertPromotable,
} = require('../../../scripts/update-feed-utils.cjs')

/** Promote-workflow guard: stable may only move forward (unless forced). */
describe('update-feed-utils', () => {
  it('parses the version from an electron-updater feed', () => {
    expect(ymlVersion('version: 0.5.82\nfiles:\n  - url: x.zip')).toBe('0.5.82')
    expect(ymlVersion('files:\n  - url: x.zip')).toBeNull()
  })

  it('compares plain x.y.z versions', () => {
    expect(semverNewer('0.5.83', '0.5.82')).toBe(true)
    expect(semverNewer('0.5.82', '0.5.82')).toBe(false)
    expect(semverNewer('0.5.9', '0.5.82')).toBe(false)
    expect(semverNewer('0.6.0', '0.5.82')).toBe(true)
  })

  it('allows promoting a strictly newer version', () => {
    expect(assertPromotable('0.5.83', '0.5.82', false)).toEqual({ ok: true })
  })

  it('allows the first promote when no stable feed exists yet', () => {
    expect(assertPromotable('0.5.83', null, false)).toEqual({ ok: true })
  })

  it('rejects equal or older versions without --force', () => {
    expect(assertPromotable('0.5.82', '0.5.82', false).ok).toBe(false)
    expect(assertPromotable('0.5.80', '0.5.82', false).ok).toBe(false)
  })

  it('lets --force roll back', () => {
    expect(assertPromotable('0.5.80', '0.5.82', true)).toEqual({ ok: true })
  })
})
