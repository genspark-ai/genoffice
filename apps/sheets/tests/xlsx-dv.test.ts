import { describe, expect, it } from 'vitest'

import { applyDvRules, DvEditError } from '../src/gateway/xlsx-dv'

const SHEET =
  '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  '<dataValidations count="1"><dataValidation type="whole" operator="equal" sqref="B1">' +
  '<formula1>7</formula1></dataValidation></dataValidations>' +
  '<hyperlinks><hyperlink ref="G1" r:id="rId1"/></hyperlinks>' +
  '<pageMargins left="0.7"/></worksheet>'

const range = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }

describe('applyDvRules', () => {
  it('replaces the existing section with the snapshot, before hyperlinks', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: {
          type: 'decimal',
          operator: 'greaterThan',
          formula1: '3',
          allowBlank: true,
        },
      },
    ])
    expect(xml).not.toContain('type="whole"')
    expect(xml).toContain(
      '<dataValidations count="1"><dataValidation type="decimal" operator="greaterThan" ' +
        'allowBlank="1" sqref="A1:A5"><formula1>3</formula1></dataValidation>' +
        '</dataValidations><hyperlinks',
    )
  })

  it('removes the section when the snapshot is empty', () => {
    const xml = applyDvRules(SHEET, [])
    expect(xml).not.toContain('dataValidations')
    expect(xml).toContain('<hyperlinks>')
  })

  it('quotes list literals and strips = from list references', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range], rule: { type: 'list', formula1: 'Yes,No', showDropDown: true } },
      {
        ranges: [{ ...range, startColumn: 1, endColumn: 1 }],
        rule: { type: 'list', formula1: "='My Sheet'!$D$1:$D$9", showDropDown: false },
      },
    ])
    expect(xml).toContain('<formula1>"Yes,No"</formula1>')
    expect(xml).toContain("<formula1>'My Sheet'!$D$1:$D$9</formula1>")
    // OOXML showDropDown="1" means "suppress the dropdown" — only the
    // dropdown-off rule carries it.
    expect(xml).toContain('<dataValidation type="list" showDropDown="1" sqref="B1:B5">')
    expect(xml).toContain('<dataValidation type="list" sqref="A1:A5">')
  })

  it('writes prompts, errors, and the inverse error-style mapping', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: {
          type: 'textLength',
          operator: 'lessThanOrEqual',
          formula1: '10',
          showInputMessage: true,
          promptTitle: 'Limit',
          prompt: 'Max 10 & "short"',
          showErrorMessage: true,
          errorStyle: 2,
          errorTitle: 'Too long',
          error: 'Shorten it',
        },
      },
    ])
    expect(xml).toContain(
      'type="textLength" operator="lessThanOrEqual" showInputMessage="1" ' +
        'showErrorMessage="1" errorStyle="warning" errorTitle="Too long" error="Shorten it" ' +
        'promptTitle="Limit" prompt="Max 10 &amp; &quot;short&quot;" sqref="A1:A5"',
    )
  })

  it('maps any back to the attribute-less default type and skips between', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: { type: 'any', showInputMessage: true, prompt: 'Read me' },
      },
      {
        ranges: [{ ...range, startColumn: 2, endColumn: 2 }],
        rule: { type: 'whole', operator: 'between', formula1: '1', formula2: '9' },
      },
    ])
    expect(xml).toContain('<dataValidation showInputMessage="1" prompt="Read me" sqref="A1:A5"/>')
    expect(xml).toContain(
      '<dataValidation type="whole" sqref="C1:C5"><formula1>1</formula1>' +
        '<formula2>9</formula2></dataValidation>',
    )
  })

  it('converts panel date/time strings to Excel serials, keeps numbers', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: {
          type: 'date',
          operator: 'notBetween',
          formula1: '2024-01-01',
          formula2: '2024-12-31 12:00:00',
        },
      },
      {
        ranges: [{ ...range, startColumn: 3, endColumn: 3 }],
        rule: { type: 'time', operator: 'lessThan', formula1: '06:00' },
      },
      {
        ranges: [{ ...range, startColumn: 4, endColumn: 4 }],
        rule: { type: 'date', operator: 'greaterThan', formula1: '45123' },
      },
    ])
    expect(xml).toContain('<formula1>45292</formula1>')
    expect(xml).toContain('<formula2>45657.5</formula2>')
    expect(xml).toContain('<formula1>0.25</formula1>')
    expect(xml).toContain('<formula1>45123</formula1>')
  })

  it('strips = from custom formulas and escapes their XML', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [range],
        rule: { type: 'custom', formula1: '=AND($A1>0,$A1<5)' },
      },
    ])
    expect(xml).toContain('<formula1>AND($A1&gt;0,$A1&lt;5)</formula1>')
  })

  it('degrades checkbox rules to a two-value list', () => {
    const xml = applyDvRules(SHEET, [
      { ranges: [range], rule: { type: 'checkbox', allowBlank: true } },
      {
        ranges: [{ ...range, startColumn: 1, endColumn: 1 }],
        rule: { type: 'checkbox', formula1: 'Yes', formula2: 'No' },
      },
    ])
    expect(xml).toContain(
      '<dataValidation type="list" allowBlank="1" sqref="A1:A5">' +
        '<formula1>"1,0"</formula1></dataValidation>',
    )
    expect(xml).toContain(
      '<dataValidation type="list" sqref="B1:B5"><formula1>"Yes,No"</formula1></dataValidation>',
    )
  })

  it('fails closed on Univer-only rule types', () => {
    expect(() =>
      applyDvRules(SHEET, [{ ranges: [range], rule: { type: 'listMultiple', formula1: 'a,b' } }]),
    ).toThrow(DvEditError)
    expect(() =>
      applyDvRules(SHEET, [{ ranges: [range], rule: { type: 'listMultiple', formula1: 'a,b' } }]),
    ).toThrow(/Multi-select/)
  })

  it('fails closed on x14 extended validations', () => {
    const x14Sheet = SHEET.replace(
      '<pageMargins left="0.7"/>',
      '<extLst><ext><x14:dataValidations count="1"><x14:dataValidation type="list"/>' +
        '</x14:dataValidations></ext></extLst>',
    )
    expect(() =>
      applyDvRules(x14Sheet, [
        { ranges: [range], rule: { type: 'whole', operator: 'equal', formula1: '1' } },
      ]),
    ).toThrow(/x14/)
  })

  it('appends before </worksheet> when no later section exists', () => {
    const bare = '<worksheet><sheetData/></worksheet>'
    const xml = applyDvRules(bare, [
      { ranges: [range], rule: { type: 'whole', operator: 'equal', formula1: '1' } },
    ])
    expect(xml).toBe(
      '<worksheet><sheetData/><dataValidations count="1">' +
        '<dataValidation type="whole" operator="equal" sqref="A1:A5">' +
        '<formula1>1</formula1></dataValidation></dataValidations></worksheet>',
    )
  })
})
