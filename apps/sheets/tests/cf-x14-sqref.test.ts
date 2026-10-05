import { describe, expect, it } from 'vitest'

import { applyCfRules } from '@genoffice/xlsx-gateway/gateway/xlsx-cf'

class FakeDxfs {
  internDxf(): number {
    return 0
  }
}

const DOLLAR_BAR =
  '<conditionalFormatting sqref="$C$1:$C$5">' +
  '<cfRule type="dataBar" priority="2">' +
  '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>' +
  '<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"' +
  ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:id>{DATA-BAR-GUID}</x14:id></ext></extLst>' +
  '</cfRule></conditionalFormatting>'

const SHEET =
  '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  `${DOLLAR_BAR}<pageMargins left="0.7"/></worksheet>`

describe('cf x14 sqref normalize', () => {
  it('matches dollarized sqref to the unchanged rule', () => {
    const xml = applyCfRules(
      SHEET,
      [
        {
          ranges: [{ startRow: 0, endRow: 4, startColumn: 2, endColumn: 2 }],
          stopIfTrue: false,
          rule: {
            type: 'dataBar',
            config: { min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6' },
          },
        },
      ],
      new FakeDxfs(),
    )
    expect(xml).toContain(DOLLAR_BAR)
  })
})
