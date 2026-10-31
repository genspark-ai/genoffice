import { describe, expect, it } from 'vitest'

import {
  a1RangeToTarget,
  boundConnections,
  queryTableBinding,
  refreshableConnections,
  tableToCellEdits,
} from '../src/gateway/xlsx-connection-refresh'

const CONNECTIONS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <connection id="1" name="Rates" type="5" refreshedVersion="8">
    <webPr url="https://example.com/rates.html"/>
  </connection>
  <connection id="2" name="Local" type="6">
    <textPr sourceFile="rates.csv"/>
  </connection>
  <connection id="3" name="Ledger" type="1">
    <dbPr connection="Provider=SQLOLEDB;Data Source=secret-host;User ID=admin;Password=hunter2"/>
  </connection>
  <connection id="4" name="Mashup" type="5">
    <dbPr connection="Provider=Microsoft.Mashup.OleDb;Data Source=$Workbook$;Location=Rates"/>
  </connection>
</connections>`

describe('refreshableConnections', () => {
  it('lists the web and text connections with what they read', () => {
    expect(refreshableConnections(CONNECTIONS_XML)).toEqual([
      { id: '1', name: 'Rates', kind: 'web', source: 'https://example.com/rates.html' },
      { id: '2', name: 'Local', kind: 'text', source: 'rates.csv' },
    ])
  })

  it('does not list a database connection, whose string can carry credentials', () => {
    const listed = refreshableConnections(CONNECTIONS_XML)
    expect(listed.some((c) => c.id === '3')).toBe(false)
    expect(JSON.stringify(listed)).not.toContain('hunter2')
  })

  it('does not list a Power Query connection, which needs an M interpreter', () => {
    expect(refreshableConnections(CONNECTIONS_XML).some((c) => c.id === '4')).toBe(false)
  })

  it('finds nothing in a package without connections', () => {
    expect(refreshableConnections('<connections/>')).toEqual([])
  })
})

describe('queryTableBinding', () => {
  it('reads the connection and the range the query table fills', () => {
    const xml =
      '<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'name="Rates" connectionId="1" ref="A1:D21"/>'
    expect(queryTableBinding(xml)).toEqual({ connectionId: '1', ref: 'A1:D21' })
  })

  it('returns null when the part names neither a connection nor a range', () => {
    expect(queryTableBinding('<queryTable name="Rates"/>')).toBeNull()
    expect(queryTableBinding('<not-a-query-table/>')).toBeNull()
  })
})

const WORKBOOK_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Rates" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

const WORKBOOK_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`

const SHEET_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/>
</Relationships>`

const QUERY_TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Rates" connectionId="1" ref="A1:D21"/>`

describe('boundConnections', () => {
  const parts = {
    connectionsXml: CONNECTIONS_XML,
    workbookXml: WORKBOOK_XML,
    workbookRelsXml: WORKBOOK_RELS_XML,
    sheetRels: [{ partPath: 'xl/worksheets/sheet1.xml', xml: SHEET_RELS_XML }],
    queryTables: [{ path: 'xl/queryTables/queryTable1.xml', xml: QUERY_TABLE_XML }],
  }

  it('binds a refreshable query to the sheet and range it fills', () => {
    expect(boundConnections(parts)).toEqual([
      {
        connection: {
          id: '1',
          name: 'Rates',
          kind: 'web',
          source: 'https://example.com/rates.html',
        },
        sheetName: 'Rates',
        range: 'A1:D21',
      },
    ])
  })

  it('leaves out a connection that has no query table to write into', () => {
    const result = boundConnections({
      ...parts,
      queryTables: [],
    })
    expect(result).toEqual([])
  })

  it('leaves out a query table whose connection is not refreshable', () => {
    const result = boundConnections({
      ...parts,
      queryTables: [
        {
          path: 'xl/queryTables/queryTable1.xml',
          xml: QUERY_TABLE_XML.replace('connectionId="1"', 'connectionId="3"'),
        },
      ],
    })
    expect(result).toEqual([])
  })

  it('finds nothing in a package without connections', () => {
    expect(boundConnections({ ...parts, connectionsXml: '<connections/>' })).toEqual([])
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
})

describe('tableToCellEdits', () => {
  const target = {
    sheetName: 'Rates',
    startRow: 0,
    startColumn: 0,
    width: 2,
    height: 3,
  }

  it('writes the header and the rows from the top-left corner', () => {
    const edits = tableToCellEdits(target, {
      headers: ['Currency', 'Rate'],
      rows: [
        ['EUR', '1.09'],
        ['USD', '1.00'],
      ],
    })
    expect(edits.map((e) => [e.row, e.column, e.cell.value])).toEqual([
      [0, 0, 'Currency'],
      [0, 1, 'Rate'],
      [1, 0, 'EUR'],
      [1, 1, '1.09'],
      [2, 0, 'USD'],
      [2, 1, '1.00'],
    ])
    expect(edits.every((e) => e.sheetName === 'Rates' && e.writeValue)).toBe(true)
  })

  it('keeps the edits inside the target range', () => {
    const edits = tableToCellEdits(target, {
      headers: ['a', 'b', 'c'],
      rows: [
        ['1', '2', '3'],
        ['4', '5', '6'],
        ['7', '8', '9'],
        ['10', '11', '12'],
      ],
    })
    for (const edit of edits) {
      expect(edit.row).toBeLessThan(target.height)
      expect(edit.column).toBeLessThan(target.width)
    }
  })

  it('writes into a target that does not start at A1', () => {
    const edits = tableToCellEdits(
      { sheetName: 'Rates', startRow: 4, startColumn: 2, width: 2, height: 2 },
      { headers: ['h1', 'h2'], rows: [['v1', 'v2']] },
    )
    expect(edits.map((e) => [e.row, e.column, e.cell.value])).toEqual([
      [4, 2, 'h1'],
      [4, 3, 'h2'],
      [5, 2, 'v1'],
      [5, 3, 'v2'],
    ])
  })

  it('writes an empty value as an empty cell', () => {
    const edits = tableToCellEdits(target, { headers: ['a'], rows: [['']] })
    expect(edits[1]?.cell.value).toBe('')
  })
})
