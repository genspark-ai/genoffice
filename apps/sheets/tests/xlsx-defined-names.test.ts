import { describe, expect, it } from 'vitest'

import { applyDefinedNamesState, DefinedNameError } from '../src/gateway/xlsx-defined-names'

const WORKBOOK =
  '<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
  '<definedNames>' +
  '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$C$4</definedName>' +
  '<definedName name="OldName">Data!$B$2</definedName>' +
  '<definedName name="Unparseable">EXOTIC(1)</definedName>' +
  '</definedNames><calcPr/></workbook>'

describe('applyDefinedNamesState', () => {
  it('rewrites modeled names, preserving built-ins and preserve-listed ones', () => {
    const xml = applyDefinedNamesState(WORKBOOK, {
      names: [
        { name: 'Revenue', formula: 'Data!$B$2:$B$9' },
        { name: 'LocalTotal', formula: 'Data!$C$1', sheetIndex: 0 },
      ],
      preserveNames: ['Unparseable'],
    })
    expect(xml).toContain('name="_xlnm._FilterDatabase"')
    expect(xml).toContain('<definedName name="Unparseable">EXOTIC(1)</definedName>')
    expect(xml).not.toContain('OldName')
    expect(xml).toContain('<definedName name="Revenue">Data!$B$2:$B$9</definedName>')
    expect(xml).toContain('<definedName name="LocalTotal" localSheetId="0">Data!$C$1</definedName>')
  })

  it('removes the section when nothing remains', () => {
    const bare =
      '<workbook><sheets><sheet name="D"/></sheets>' +
      '<definedNames><definedName name="Gone">D!$A$1</definedName></definedNames></workbook>'
    expect(applyDefinedNamesState(bare, { names: [], preserveNames: [] })).toBe(
      '<workbook><sheets><sheet name="D"/></sheets></workbook>',
    )
  })

  it('creates the section after sheets when absent, stripping a leading =', () => {
    const bare = '<workbook><sheets><sheet name="D"/></sheets><calcPr/></workbook>'
    expect(
      applyDefinedNamesState(bare, {
        names: [{ name: 'N', formula: '=D!$A$1' }],
        preserveNames: [],
      }),
    ).toBe(
      '<workbook><sheets><sheet name="D"/></sheets>' +
        '<definedNames><definedName name="N">D!$A$1</definedName></definedNames>' +
        '<calcPr/></workbook>',
    )
  })

  it('escapes XML in formulas', () => {
    const xml = applyDefinedNamesState(WORKBOOK, {
      names: [{ name: 'Cmp', formula: 'IF(A1<5,"a & b","c")' }],
      preserveNames: [],
    })
    expect(xml).toContain('<definedName name="Cmp">IF(A1&lt;5,"a &amp; b","c")</definedName>')
  })

  it('accepts Unicode names the way Excel does', () => {
    // Create from Selection derives names from localized headers (like the
    // CJK header below); rejecting them here would fail the whole workbook save.
    const xml = applyDefinedNamesState(WORKBOOK, {
      names: [
        { name: '销售额', formula: 'Data!$B$2:$B$9' },
        { name: 'Umsätze_2024', formula: 'Data!$C$2:$C$9' },
        { name: '_売上', formula: 'Data!$D$2:$D$9' },
      ],
      preserveNames: [],
    })
    expect(xml).toContain('<definedName name="销售额">Data!$B$2:$B$9</definedName>')
    expect(xml).toContain('<definedName name="Umsätze_2024">Data!$C$2:$C$9</definedName>')
    expect(xml).toContain('<definedName name="_売上">Data!$D$2:$D$9</definedName>')
  })

  it('fails closed on invalid, reserved, duplicate, or conflicting names', () => {
    const save = (name: string) => () =>
      applyDefinedNamesState(WORKBOOK, {
        names: [{ name, formula: 'Data!$A$1' }],
        preserveNames: [],
      })
    expect(save('has space')).toThrow(DefinedNameError)
    expect(save('A1')).toThrow(DefinedNameError)
    expect(save('R1C1')).toThrow(DefinedNameError)
    expect(save('true')).toThrow(DefinedNameError)
    expect(save('_xlnm.Custom')).toThrow(DefinedNameError)
    expect(() =>
      applyDefinedNamesState(WORKBOOK, {
        names: [
          { name: 'Twice', formula: 'Data!$A$1' },
          { name: 'Twice', formula: 'Data!$A$2' },
        ],
        preserveNames: [],
      }),
    ).toThrow(/defined twice/)
    expect(() =>
      applyDefinedNamesState(WORKBOOK, {
        names: [{ name: 'Unparseable', formula: 'Data!$A$1' }],
        preserveNames: ['Unparseable'],
      }),
    ).toThrow(/duplicate/)
  })

  it('allows the same name in different scopes', () => {
    const xml = applyDefinedNamesState(WORKBOOK, {
      names: [
        { name: 'Total', formula: 'Data!$A$1' },
        { name: 'Total', formula: 'Data!$B$1', sheetIndex: 0 },
      ],
      preserveNames: [],
    })
    expect(xml).toContain('<definedName name="Total">Data!$A$1</definedName>')
    expect(xml).toContain('<definedName name="Total" localSheetId="0">Data!$B$1</definedName>')
  })
})
