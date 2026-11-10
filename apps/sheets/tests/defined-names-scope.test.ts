import { describe, expect, it } from 'vitest'

import {
  applyDefinedNamesState,
  DefinedNameError,
} from '@genoffice/xlsx-gateway/gateway/xlsx-defined-names'

const WORKBOOK =
  '<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
  '<definedNames><definedName name="MyName">Data!$A$1</definedName></definedNames>' +
  '<calcPr/></workbook>'

describe('defined-names scope', () => {
  it('allows a preserved workbook name alongside the same sheet-scoped name', () => {
    const xml = applyDefinedNamesState(WORKBOOK, {
      names: [{ name: 'MyName', formula: 'Data!$B$1', sheetIndex: 0 }],
      preserveNames: ['MyName'],
    })
    expect(xml).toContain('<definedName name="MyName">Data!$A$1</definedName>')
    expect(xml).toContain('<definedName name="MyName" localSheetId="0">Data!$B$1</definedName>')
  })

  it('still rejects a preserved name in the same scope', () => {
    expect(() =>
      applyDefinedNamesState(WORKBOOK, {
        names: [{ name: 'MyName', formula: 'Data!$B$1' }],
        preserveNames: ['MyName'],
      }),
    ).toThrow(DefinedNameError)
  })
})
