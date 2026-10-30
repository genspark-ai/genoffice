import { describe, expect, it } from 'vitest'
import { listWorkbookQueries, parseConnections } from '../src/gateway/xlsx-connections'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/// The parts Excel writes for a Power Query loaded to a table on "Rates", a
/// connection-only query, a legacy web query and a text import.
const parts: Record<string, string> = {
  'xl/connections.xml':
    '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<connection id="1" keepAlive="1" name="Query - Rates" description="Connection to the &apos;Rates&apos; query in the workbook." type="5" refreshedVersion="8" background="1" saveData="1">' +
    '<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Rates;Extended Properties=&quot;&quot;" command="SELECT * FROM [Rates]"/></connection>' +
    '<connection id="2" keepAlive="1" name="Query - Sales 2024" type="5" refreshedVersion="8" background="1">' +
    '<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=&quot;Sales 2024&quot;;Extended Properties=&quot;&quot;" command="SELECT * FROM [Sales 2024]"/></connection>' +
    '<connection id="3" name="Connection" type="4" refreshedVersion="2"><webPr url="https://example.com/rates.html" htmlTables="1"/></connection>' +
    '<connection id="4" name="prices" type="6" refreshedVersion="8"><textPr sourceFile="C:\\data\\prices.csv" delimited="1"/></connection>' +
    '</connections>',
  'xl/workbook.xml': `<workbook xmlns:r="${REL}"><sheets><sheet name="Rates" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets></workbook>`,
  'xl/_rels/workbook.xml.rels':
    '<Relationships>' +
    `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
    `<Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>` +
    '</Relationships>',
  'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`,
  'xl/tables/table1.xml':
    '<table id="1" name="Rates" displayName="Rates" ref="A1:C21" tableType="queryTable"/>',
  'xl/tables/_rels/table1.xml.rels': `<Relationships><Relationship Id="rId1" Type="${REL}/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>`,
  'xl/queryTables/queryTable1.xml':
    '<queryTable name="ExternalData_1" connectionId="1" autoFormatId="16"/>',
}

const reader = (files: Record<string, string>) => async (path: string) => files[path]

describe('parseConnections', () => {
  it('names Power Query connections by their query, quoted or not', () => {
    const names = parseConnections(parts['xl/connections.xml']!).map((c) => c.query.name)
    expect(names).toEqual(['Rates', 'Sales 2024', 'Connection', 'prices'])
  })
})

describe('listWorkbookQueries', () => {
  it('lists each connection with its kind, source and where it loads', async () => {
    expect(await listWorkbookQueries(reader(parts))).toEqual([
      {
        name: 'Rates',
        kind: 'powerQuery',
        description: "Connection to the 'Rates' query in the workbook.",
        destination: { sheet: 'Rates', table: 'Rates', range: 'A1:C21', rows: 20 },
      },
      { name: 'Sales 2024', kind: 'powerQuery' },
      { name: 'Connection', kind: 'web', source: 'https://example.com/rates.html' },
      { name: 'prices', kind: 'text', source: 'C:\\data\\prices.csv' },
    ])
  })

  it("counts data rows without a table's header and totals rows", async () => {
    const totals = {
      ...parts,
      'xl/tables/table1.xml':
        '<table id="1" displayName="Rates" ref="A1:C21" totalsRowCount="1" tableType="queryTable"/>',
    }
    expect((await listWorkbookQueries(reader(totals)))[0]?.destination?.rows).toBe(19)
    const headless = {
      ...parts,
      'xl/tables/table1.xml':
        '<table id="1" displayName="Rates" ref="A1:C21" headerRowCount="0" tableType="queryTable"/>',
    }
    expect((await listWorkbookQueries(reader(headless)))[0]?.destination?.rows).toBe(21)
  })

  it('finds a legacy query table loaded straight onto a sheet by its defined name', async () => {
    const legacy = {
      ...parts,
      'xl/workbook.xml': parts['xl/workbook.xml']!.replace(
        '</sheets>',
        '</sheets><definedNames>' +
          '<definedName name="ExternalData_1" localSheetId="0" hidden="1">Rates!$A$1:$B$4</definedName>' +
          '<definedName name="ExternalData_1" localSheetId="1" hidden="1">Notes!$D$2:$E$12</definedName>' +
          '</definedNames>',
      ),
      'xl/worksheets/_rels/sheet2.xml.rels': `<Relationships><Relationship Id="rId1" Type="${REL}/queryTable" Target="../queryTables/queryTable2.xml"/></Relationships>`,
      'xl/queryTables/queryTable2.xml': '<queryTable name="ExternalData_1" connectionId="3"/>',
    }
    expect((await listWorkbookQueries(reader(legacy)))[2]?.destination).toEqual({
      sheet: 'Notes',
      range: 'D2:E12',
      rows: 10,
    })
  })

  it('clips what the result schema would refuse, so one entry cannot fail the list', async () => {
    const huge = {
      'xl/connections.xml':
        `<connections><connection id="1" name="${'n'.repeat(2_000)}" description="${'d'.repeat(9_000)}">` +
        `<webPr url="https://example.com/${'u'.repeat(9_000)}"/></connection>` +
        Array.from({ length: 1_500 }, (_, i) => `<connection id="${i + 2}" name="c"/>`).join('') +
        '</connections>',
    }
    const queries = await listWorkbookQueries(reader(huge))
    expect(queries).toHaveLength(1_000)
    expect(queries[0]?.name).toHaveLength(1_000)
    expect(queries[0]?.source).toHaveLength(4_000)
    expect(queries[0]?.description).toHaveLength(4_000)
  })

  it('is empty for a workbook without connections', async () => {
    const { ['xl/connections.xml']: _omit, ...rest } = parts
    expect(await listWorkbookQueries(reader(rest))).toEqual([])
  })

  it('skips a broken table link instead of failing the whole list', async () => {
    const broken = {
      ...parts,
      'xl/tables/_rels/table1.xml.rels':
        '<Relationships><Relationship Id="rId1" Type="x/queryTable" Target="http://evil/"/></Relationships>',
    }
    const queries = await listWorkbookQueries(reader(broken))
    expect(queries).toHaveLength(4)
    expect(queries[0]?.destination).toBeUndefined()
  })
})
