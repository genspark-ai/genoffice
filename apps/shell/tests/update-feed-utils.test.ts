import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'

// the utils are a dependency-free CommonJS script shared with CI; load via
// createRequire since this test file is ESM
const require = createRequire(import.meta.url)
const {
  ymlVersion,
  semverNewer,
  assertPromotable,
  releaseUploadDecision,
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

  it('makes an exact-version CI rerun a no-op when explicitly allowed', () => {
    expect(
      releaseUploadDecision('0.5.82', '0.5.82', {
        allowExisting: true,
      }),
    ).toEqual({ action: 'skip' })
  })

  it('still rejects equal and older uploads by default', () => {
    expect(releaseUploadDecision('0.5.82', '0.5.82').action).toBe('reject')
    expect(releaseUploadDecision('0.5.81', '0.5.82', { allowExisting: true }).action).toBe('reject')
  })

  describe('prerelease ordering', () => {
    const table: Array<[string, string, boolean]> = [
      ['0.11.0-beta.2', '0.11.0-beta.1', true],
      ['0.11.0-beta.10', '0.11.0-beta.9', true],
      ['0.11.0-beta.1', '0.11.0-beta.2', false],
      ['0.11.0-beta.2', '0.11.0-beta.2', false],
      ['0.11.0', '0.11.0-beta.9', true],
      ['0.11.0-beta.9', '0.11.0', false],
      ['0.11.0-beta', '0.11.0-alpha', true],
      ['0.11.0-beta', '0.11.0-beta.1', false],
      ['0.11.0-alpha.beta', '0.11.0-alpha.1', true],
      ['0.11.0-alpha.1', '0.11.0-alpha.beta', false],
      ['0.11.0-beta', '0.11.0-alpha.beta', true],
      ['0.11.0-rc.1', '0.11.0-beta.9', true],
      ['0.11.1-beta.1', '0.11.0', true],
      ['0.11.0+build.2', '0.11.0+build.1', false],
      ['0.11.0+build.2', '0.11.0-beta.1', true],
    ]

    it.each(table)('semverNewer(%s, %s) === %s', (a, b, expected) => {
      expect(semverNewer(a, b)).toBe(expected)
    })

    it('never reports an unparseable version as newer', () => {
      expect(semverNewer('beta', '0.5.82')).toBe(false)
      expect(semverNewer('0.5.83', 'not-a-version')).toBe(false)
    })

    it('uploads a newer prerelease and rejects going backwards', () => {
      expect(releaseUploadDecision('0.11.0-beta.2', '0.11.0-beta.1')).toEqual({ action: 'upload' })
      expect(releaseUploadDecision('0.11.0-beta.1', '0.11.0-beta.2').action).toBe('reject')
      expect(releaseUploadDecision('0.11.0', '0.11.0-beta.2')).toEqual({ action: 'upload' })
    })

    it('promotes a newer prerelease onto the stable feed', () => {
      expect(assertPromotable('0.11.0-beta.2', '0.11.0-beta.1', false)).toEqual({ ok: true })
      expect(assertPromotable('0.11.0', '0.11.0-beta.2', false)).toEqual({ ok: true })
      expect(assertPromotable('0.11.0-beta.1', '0.11.0-beta.2', false).ok).toBe(false)
    })
  })
})
