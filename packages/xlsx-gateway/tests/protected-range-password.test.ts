/**
 * Allow-edit ranges with a per-range password: the writer emits the same
 * password attribute shapes sheet protection uses (legacy 4-hex and the
 * spi algorithmName/hashValue/saltValue/spinCount set), and the fail-closed
 * guards stay — a rewrite still refuses when the file's own ranges carry a
 * password or a securityDescriptor.
 */
import { describe, expect, it } from 'vitest'
import { applyProtectedRanges, type ProtectedRangeState } from '../src/gateway/xlsx-protection'

const SHEET = '<sheetData/>'
const withRanges = (ranges: readonly ProtectedRangeState[]): string =>
  applyProtectedRanges(SHEET, ranges)

describe('protectedRange passwords', () => {
  it('writes a legacy hash as the password attribute', () => {
    const xml = withRanges([{ name: 'R1', sqref: 'A1:B4', password: { legacy: '8492' } }])
    expect(xml).toContain('<protectedRange sqref="A1:B4" name="R1" password="8492"/>')
  })

  it('writes the spi attribute set for a modern hash', () => {
    const xml = withRanges([
      {
        name: 'R1',
        sqref: 'A1:B4',
        password: {
          algorithmName: 'SHA-512',
          hashValue: 'QUJD',
          saltValue: 'U0FMVA==',
          spinCount: 100000,
        },
      },
    ])
    expect(xml).toContain('algorithmName="SHA-512"')
    expect(xml).toContain('hashValue="QUJD"')
    expect(xml).toContain('saltValue="U0FMVA=="')
    expect(xml).toContain('spinCount="100000"')
  })

  it('omits the attribute when no password is set', () => {
    const xml = withRanges([{ name: 'R1', sqref: 'A1:B4' }])
    expect(xml).toContain('<protectedRange sqref="A1:B4" name="R1"/>')
    expect(xml).not.toContain('password=')
  })

  it('still refuses a rewrite of ranges carrying a password or an ACL', () => {
    const guarded =
      '<sheetData/><protectedRanges><protectedRange name="R" sqref="A1" password="8492"/></protectedRanges>'
    expect(() => applyProtectedRanges(guarded, [{ name: 'R2', sqref: 'C1' }])).toThrow(
      /not.*supported|password- or permission-protected/,
    )
    const acl =
      '<sheetData/><protectedRanges><protectedRange name="R" sqref="A1" securityDescriptor="O:S-1"/></protectedRanges>'
    expect(() => applyProtectedRanges(acl, [{ name: 'R2', sqref: 'C1' }])).toThrow()
  })
})
