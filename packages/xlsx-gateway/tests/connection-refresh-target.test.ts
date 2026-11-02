import { describe, expect, it } from 'vitest'

import {
  a1RangeToTarget,
  connectionSourceOf,
  planTableRefresh,
  refreshableQueriesOf,
} from '../src/gateway/xlsx-connection-refresh'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/// The parts Excel writes for an Excel-authored workbook: a Power Query loaded
/// to a table, a ListObject-backed web query, a text query, a legacy web query
/// loaded straight onto a sheet (no table part, only ExternalData_1), and a
/// connection-only web query. "Sales &amp; Notes" keeps an escaped sheet name
/// on the legacy route.
const PARTS: Record<string, string> = {
  'xl/connections.xml':
    '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<connection id="1" name="Query - Rates" type="5">' +
    '<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Rates"/></connection>' +
    '<connection id="2" name="Rates" type="4">' +
    '<webPr url="https://example.com/rates.html"/></connection>' +
    '<connection id="3" name="prices" type="6">' +
    '<textPr sourceFile="C:\\data\\prices.csv"/></connection>' +
    '<connection id="4" name="Legacy" type="4">' +
    '<webPr url="https://example.com/legacy.html"/></connection>' +
    '<connection id="5" name="No destination" type="4">' +
    '<webPr url="https://example.com/none.html"/></connection>' +
    '</connections>',
  'xl/workbook.xml':
    '<workbook xmlns:r="' +
    REL +
    '"><sheets><sheet name="Rates" sheetId="1" r:id="rId1"/>' +
    '<sheet name="Sales &amp; Notes" sheetId="2" r:id="rId2"/></sheets>' +
    '<definedNames><definedName name="ExternalData_1" localSheetId="1" hidden="1">' +
    "'Sales &amp; Notes'!$D$2:$E$12</definedName></definedNames></workbook>",
  'xl/_rels/workbook.xml.rels':
    '<Relationships>' +
    '<Relationship Id="rId1" Type="' +
    REL +
    '/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="' +
    REL +
    '/worksheet" Target="worksheets/sheet2.xml"/>' +
    '</Relationships>',
  'xl/worksheets/_rels/sheet1.xml.rels':
    '<Relationships><Relationship Id="rId1" Type="' +
    REL +
    '/table" Target="../tables/table1.xml"/></Relationships>',
  'xl/tables/table1.xml':
    '<table id="1" name="Rates" displayName="Rates" ref="A1:C21" tableType="queryTable"/>',
  'xl/tables/_rels/table1.xml.rels':
    '<Relationships><Relationship Id="rId1" Type="' +
    REL +
    '/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>',
  'xl/queryTables/queryTable1.xml':
    '<queryTable name="ExternalData_1" connectionId="2" autoFormatId="16"/>',
  'xl/worksheets/_rels/sheet2.xml.rels':
    '<Relationships><Relationship Id="rId1" Type="' +
    REL +
    '/queryTable" Target="../queryTables/queryTable2.xml"/></Relationships>',
  'xl/queryTables/queryTable2.xml': '<queryTable name="ExternalData_1" connectionId="4"/>',
}

const reader =
  (files: Record<string, string>) =>
  async (path: string): Promise<string | undefined> =>
    files[path]

describe('refreshableQueriesOf', () => {
  it('finds a ListObject-backed web query through the table part Excel writes', async () => {
    expect(await refreshableQueriesOf(reader(PARTS))).toEqual([
      {
        id: '2',
        name: 'Rates',
        kind: 'web',
        source: 'https://example.com/rates.html',
        sheetName: 'Rates',
        range: 'A1:C21',
      },
      {
        id: '4',
        name: 'Legacy',
        kind: 'web',
        source: 'https://example.com/legacy.html',
        sheetName: 'Sales & Notes',
        range: 'D2:E12',
      },
    ])
  })

  it('decodes the escaped sheet name from the legacy defined-name route', async () => {
    const queries = await refreshableQueriesOf(reader(PARTS))
    expect(queries[1]?.sheetName).toBe('Sales & Notes')
    expect(queries[1]?.range).toBe('D2:E12')
  })

  it('leaves out a text query, a Power Query and a connection with no destination', async () => {
    const sources = (await refreshableQueriesOf(reader(PARTS))).map((query) => query.id)
    expect(sources).toEqual(['2', '4'])
  })

  it('finds nothing in a workbook without a connections part', async () => {
    const { ['xl/connections.xml']: _omit, ...rest } = PARTS
    expect(await refreshableQueriesOf(reader(rest))).toEqual([])
  })
})

describe('connectionSourceOf', () => {
  it('resolves the URL from the connection id, not from the caller', async () => {
    expect(await connectionSourceOf(reader(PARTS), '2')).toBe('https://example.com/rates.html')
    expect(await connectionSourceOf(reader(PARTS), '4')).toBe('https://example.com/legacy.html')
  })

  it('returns null for a text connection, an unknown id or no connections part', async () => {
    expect(await connectionSourceOf(reader(PARTS), '3')).toBeNull()
    expect(await connectionSourceOf(reader(PARTS), '999')).toBeNull()
    const { ['xl/connections.xml']: _omit, ...rest } = PARTS
    expect(await connectionSourceOf(reader(rest), '2')).toBeNull()
  })
})

describe('a1RangeToTarget', () => {
  it('reads a two-column range', () => {
    expect(a1RangeToTarget('A1:D21', 'Rates')).toEqual({
      sheetName: 'Rates',
      startRow: 0,
      startColumn: 0,
      width: 4,
      height: 21,
    })
  })

  it('reads a single cell as a one-by-one range', () => {
    expect(a1RangeToTarget('B7', 'Rates')).toEqual({
      sheetName: 'Rates',
      startRow: 6,
      startColumn: 1,
      width: 1,
      height: 1,
    })
  })

  it('reads a range that does not start at A1', () => {
    const target = a1RangeToTarget('C4:E9', 'Sheet2')
    expect(target?.startRow).toBe(3)
    expect(target?.startColumn).toBe(2)
    expect(target?.width).toBe(3)
    expect(target?.height).toBe(6)
  })

  it('rejects a range that is not a range', () => {
    expect(a1RangeToTarget('not a range', 'Rates')).toBeNull()
    expect(a1RangeToTarget('', 'Rates')).toBeNull()
  })

  it('rejects a backwards range', () => {
    expect(a1RangeToTarget('D5:A1', 'Rates')).toBeNull()
  })

  it('rejects a range past the sheet grid or too large to plan', () => {
    expect(a1RangeToTarget('XFE1', 'Rates')).toBeNull() // one column past XFD
    expect(a1RangeToTarget('A1048577', 'Rates')).toBeNull() // one row past the grid
    expect(a1RangeToTarget('A1:XFD1048576', 'Rates')).toBeNull() // the whole grid
    expect(a1RangeToTarget('A1048576', 'Rates')).not.toBeNull() // the last cell is fine
  })
})

describe('planTableRefresh', () => {
  const target = { sheetName: 'Rates', startRow: 4, startColumn: 2, width: 3, height: 4 }

  it('writes the header and rows and leaves the rest of the range blank', () => {
    const plan = planTableRefresh(target, {
      headers: ['Currency', 'Rate', 'Change'],
      rows: [
        ['EUR', '1.09', '-0.02'],
        ['USD', '1.00', 'n/a'],
      ],
    })
    expect(plan.values).toEqual([
      ['Currency', 'Rate', 'Change'],
      ['EUR', 1.09, -0.02],
      ['USD', 1, 'n/a'],
      [null, null, null],
    ])
    expect(plan.droppedRows).toBe(0)
    expect(plan.droppedColumns).toBe(0)
  })

  it('covers the whole destination range, so a shrinking answer clears stale rows', () => {
    const plan = planTableRefresh(
      { sheetName: 'Rates', startRow: 0, startColumn: 0, width: 2, height: 5 },
      { headers: ['h1', 'h2'], rows: [] },
    )
    expect(plan.values).toHaveLength(5)
    expect(plan.values.every((row) => row.length === 2)).toBe(true)
    for (const row of plan.values.slice(1)) expect(row).toEqual([null, null])
  })

  it('counts the rows and columns that no longer fit the fixed range', () => {
    const plan = planTableRefresh(
      { sheetName: 'Rates', startRow: 0, startColumn: 0, width: 2, height: 3 },
      {
        headers: ['a', 'b', 'c'],
        rows: [
          ['1', '2', '3'],
          ['4', '5', '6'],
          ['7', '8', '9'],
        ],
      },
    )
    expect(plan.values).toHaveLength(3)
    expect(plan.droppedRows).toBe(1)
    expect(plan.droppedColumns).toBe(1)
  })

  it('converts plain numbers and leaves locale-formatted text alone', () => {
    const plan = planTableRefresh(
      { sheetName: 'Rates', startRow: 0, startColumn: 0, width: 8, height: 2 },
      {
        headers: ['int', 'decimal', 'negative', 'exponent', 'grouped', 'percent', 'date', 'blank'],
        rows: [[' 42 ', '.5', '-3', '1e3', '1,09', '5%', '2024-01-01', '']],
      },
    )
    expect(plan.values[1]).toEqual([42, 0.5, -3, 1000, '1,09', '5%', '2024-01-01', null])
  })

  it('pads a short row with blanks instead of shifting the row', () => {
    const plan = planTableRefresh(
      { sheetName: 'Rates', startRow: 0, startColumn: 0, width: 3, height: 2 },
      { headers: ['a', 'b', 'c'], rows: [['1'], ['2', '3']] },
    )
    expect(plan.values).toEqual([
      ['a', 'b', 'c'],
      [1, null, null],
    ])
  })
})
