import { describe, expect, it } from 'vitest'

import { formatFontBytes, offerableRows, type FontCatalogRow } from '../src/font-catalog'

/**
 * The number a reader is asked to agree to.
 *
 * A download size is the one thing in this feature that must not be rounded
 * into a promise the disk cannot keep, and it crosses no language boundary —
 * "28 MiB" is the same in every locale, so it needs no translation and cannot
 * drift from what gets written.
 */
describe('formatFontBytes', () => {
  it('reads like a person would say it', () => {
    expect(formatFontBytes(0)).toBe('0 B')
    expect(formatFontBytes(512)).toBe('512 B')
    expect(formatFontBytes(1024)).toBe('1 KiB')
    expect(formatFontBytes(1536)).toBe('2 KiB')
    expect(formatFontBytes(5 * 1024 * 1024)).toBe('5.0 MiB')
    expect(formatFontBytes(28.4 * 1024 * 1024)).toBe('28 MiB')
  })

  it('keeps a decimal only where it would be thrown away', () => {
    // 4.2 MiB rendered as "4 MiB" understates what the reader agreed to.
    expect(formatFontBytes(4.2 * 1024 * 1024)).toBe('4.2 MiB')
    // 28.4 rounds, because a decimal there is noise.
    expect(formatFontBytes(28.4 * 1024 * 1024)).toBe('28 MiB')
  })

  it('spells the binary unit out rather than saying MB', () => {
    // The catalog counts bytes; "28 MB" would be a claim about the download
    // that does not match the bytes written.
    expect(formatFontBytes(28 * 1024 * 1024)).not.toContain('MB')
  })

  it('says nothing rather than something wrong', () => {
    expect(formatFontBytes(-1)).toBe('')
    expect(formatFontBytes(Number.NaN)).toBe('')
    expect(formatFontBytes(Number.POSITIVE_INFINITY)).toBe('')
  })
})

describe('which rows a picker offers', () => {
  const row = (over: Partial<FontCatalogRow> = {}): FontCatalogRow => ({
    family: 'Noto Serif SC',
    script: 'sc',
    license: 'OFL-1.1',
    installed: false,
    bytes: 29_000_000,
    ...over,
  })

  it('offers what is not installed yet', () => {
    const rows = offerableRows([row(), row({ family: 'Roboto', installed: true })])
    expect(rows.map((r) => r.family)).toEqual(['Noto Serif SC'])
  })

  it('carries the size a row needs to ask first', () => {
    const [offer] = offerableRows([row()])
    expect(formatFontBytes(offer!.bytes)).toBe('28 MiB')
  })
})
