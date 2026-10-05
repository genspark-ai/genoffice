import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { FONT_CATALOG, type FontLicense } from '../src/main/font-catalog'

/**
 * The catalogue the app DOWNLOADS at runtime, as opposed to the fonts shipped
 * inside the bundle (those are documented in
 * apps/docs/src/renderer/fonts/README.md, and NOTICE points there).
 *
 * These bytes come off the mirror at the user's request, so the release that
 * ships the catalogue is the one that has to account for their terms. The
 * `license` field is what makes that checkable: "they are all OFL today" is
 * true, but a comment is not something a new family has to look at, and the
 * next family to be added is exactly the one that would carry a different
 * licence.
 */

/** The permissive half of tools/check-licenses.mjs's npm allowlist. */
const ALLOWED: ReadonlySet<string> = new Set<FontLicense>(['OFL-1.1', 'Apache-2.0', 'MIT'])

describe('downloadable font catalogue licensing', () => {
  it('gives every family a licence from the allowlist', () => {
    const unlicensed = FONT_CATALOG.filter((family) => !ALLOWED.has(family.license))
    expect(
      unlicensed.map((family) => `${family.family} (${family.license})`),
      'a downloadable family carries a licence outside the allowlist',
    ).toEqual([])
  })

  it('names a licence on every family, so none is silently omitted', () => {
    // `license` is required by the type, so a missing one is a compile error —
    // this asserts the runtime shape too, for a catalogue built from JSON or a
    // future generator that casts.
    for (const family of FONT_CATALOG) {
      expect(typeof family.license, family.family).toBe('string')
      expect(family.license.length, family.family).toBeGreaterThan(0)
    }
  })

  it('covers the CJK families the picker offers', () => {
    // The families #1144 added are the reason the field exists at all; a
    // check that only covered the Latin rows would pass while the interesting
    // case went unexamined.
    const cjk = FONT_CATALOG.filter((family) => family.script !== 'latin')
    expect(cjk.length).toBeGreaterThanOrEqual(9)
    for (const family of cjk) {
      expect(ALLOWED.has(family.license), family.family).toBe(true)
    }
  })

  it('lists exactly the catalogue families in NOTICE', () => {
    // A NOTICE that drifts from the catalogue is worse than none: it is the
    // document a licence audit reads. The paragraph is the one under
    // "Downloadable fonts", and the family list is the comma-separated run that
    // ends the OFL sentence.
    const notice = readFileSync(join(__dirname, '../../../NOTICE'), 'utf8')
    const paragraph = notice.slice(notice.indexOf('Downloadable fonts'))
    const listed = /families are: ([^.]+)\./.exec(paragraph)?.[1] ?? ''
    const names = listed
      .split(',')
      .map((name) => name.trim())
      .filter(Boolean)
      .sort()
    expect(names).toEqual(FONT_CATALOG.map((family) => family.family).sort())
  })

  it('agrees with the file header about what the mirror serves', () => {
    // The catalogue is generated, so its contract lives in that comment. If the
    // mirror ever starts serving a second licence, this is the place that says
    // so before the allowlist quietly grows.
    const nonOfl = FONT_CATALOG.filter((family) => family.license !== 'OFL-1.1')
    expect(nonOfl.map((family) => family.family)).toEqual([])
  })
})
