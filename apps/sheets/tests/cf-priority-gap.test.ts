import { describe, expect, it } from 'vitest'

import { applyCfRules } from '@genoffice/xlsx-gateway/gateway/xlsx-cf'

class FakeDxfs {
  internDxf(): number {
    return 0
  }
}

const GAP_SHEET =
  '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>' +
  '<conditionalFormatting sqref="B1"><cfRule type="cellIs" dxfId="0" priority="3" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>' +
  '</worksheet>'

describe('cf priority gap', () => {
  it('assigns max+1 instead of filling gaps', () => {
    const xml = applyCfRules(
      GAP_SHEET,
      [
        {
          ranges: [{ startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 }],
          stopIfTrue: false,
          rule: {
            type: 'highlightCell',
            subType: 'number',
            operator: 'greaterThan',
            value: 5,
            style: { bg: { rgb: '#FFF2CC' } },
          },
        },
      ],
      new FakeDxfs(),
      { append: true },
    )
    expect(xml).toContain('priority="4"')
    expect(xml).not.toContain('priority="2"')
  })
})
