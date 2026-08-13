import { describe, expect, it } from 'vitest'

import { applySheetProtection, SheetProtectionError } from '../src/gateway/xlsx-protection'

const BARE = '<worksheet><sheetData/><autoFilter ref="A1:C4"/></worksheet>'

describe('applySheetProtection', () => {
  it('inserts the element after sheetData with Excel defaults', () => {
    expect(applySheetProtection(BARE, true)).toBe(
      '<worksheet><sheetData/><sheetProtection sheet="1" objects="1" scenarios="1"/>' +
        '<autoFilter ref="A1:C4"/></worksheet>',
    )
  })

  it('removes an unpassworded element and is a no-op without one', () => {
    const protectedXml = applySheetProtection(BARE, true)
    expect(applySheetProtection(protectedXml, false)).toBe(BARE)
    expect(applySheetProtection(BARE, false)).toBe(BARE)
  })

  it('re-enables the sheet attribute on an existing element, keeping others', () => {
    const xml =
      '<worksheet><sheetData/>' +
      '<sheetProtection sheet="0" formatCells="0" insertRows="0"/></worksheet>'
    expect(applySheetProtection(xml, true)).toContain(
      '<sheetProtection sheet="1" formatCells="0" insertRows="0"/>',
    )
    const already = applySheetProtection(xml, true)
    expect(applySheetProtection(already, true)).toBe(already)
  })

  it('fails closed when unprotecting a password-protected sheet', () => {
    for (const attrs of [
      'sheet="1" password="83AF"',
      'sheet="1" algorithmName="SHA-512" hashValue="x" saltValue="y" spinCount="100000"',
    ]) {
      const xml = `<worksheet><sheetData/><sheetProtection ${attrs}/></worksheet>`
      expect(() => applySheetProtection(xml, false)).toThrow(SheetProtectionError)
    }
  })

  it('handles the paired-tag form', () => {
    const xml = '<worksheet><sheetData/><sheetProtection sheet="1"></sheetProtection></worksheet>'
    expect(applySheetProtection(xml, false)).toBe('<worksheet><sheetData/></worksheet>')
  })
})
