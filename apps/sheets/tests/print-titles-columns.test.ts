import { describe, expect, it } from 'vitest'

import { applyPrintAreas, PageSetupError } from '@genoffice/xlsx-gateway/gateway/xlsx-page-setup'

const WORKBOOK =
  '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'

describe('print titles columns', () => {
  it('writes column spans and combined row+column titles', () => {
    const cols = applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printTitles: 'A:A' }])
    expect(cols).toContain("'Sheet1'!$A:$A")
    const both = applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printTitles: 'A:B,1:3' }])
    expect(both).toContain("'Sheet1'!$A:$B,'Sheet1'!$1:$3")
  })

  it('keeps stored columns when setting rows', () => {
    const withCols = applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printTitles: 'A:A' }])
    const merged = applyPrintAreas(withCols, [{ sheetName: 'Sheet1', printTitles: '1:2' }])
    expect(merged).toContain("'Sheet1'!$A:$A")
    expect(merged).toContain("'Sheet1'!$1:$2")
  })

  it('rejects reversed and mixed spans', () => {
    for (const titles of ['3:1', 'B:A', 'A:3']) {
      expect(() =>
        applyPrintAreas(WORKBOOK, [{ sheetName: 'Sheet1', printTitles: titles }]),
      ).toThrow(PageSetupError)
    }
  })
})
