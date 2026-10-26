import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
  type EntrySource,
} from '../src/gateway/xlsx-gateway'
import type { FilterColumnState, SheetFilterState } from '../src/gateway/xlsx-filter'

/// Excel stores a ListObject's filter inside the table part, never on the
/// worksheet. These tests cover the save routing end to end: a filter whose
/// range sits inside a table lands in xl/tables/tableN.xml (the worksheet
/// keeps only row visibility), any other range keeps the worksheet
/// <autoFilter>, and a range straddling a table boundary fails closed.

const SHEET = 'Data'
const TABLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'

interface TableSpec {
  readonly path: string
  readonly name: string
  readonly ref: string
  /// Extra attributes on <table> (headerRowCount, totalsRowCount, ...).
  readonly attributes?: string
  /// The autoFilter slot; defaults to a bare element like Excel writes for an
  /// unfiltered table.
  readonly autoFilter?: string
}

function contentTypes(tables: readonly TableSpec[]): string {
  const overrides = tables
    .map(
      (table) =>
        `  <Override PartName="/${table.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`,
    )
    .join('\n')
  return `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
${overrides}
</Types>`
}

function worksheet(tableCount: number): string {
  const tableParts =
    tableCount === 0
      ? ''
      : `<tableParts count="${tableCount}">${Array.from(
          { length: tableCount },
          (_, index) => `<tablePart r:id="rId${index + 1}"/>`,
        ).join('')}</tableParts>`
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>10</v></c><c r="C1"><v>100</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>20</v></c><c r="C2"><v>200</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>30</v></c><c r="C3"><v>300</v></c></row>
    <row r="4"><c r="A4"><v>4</v></c><c r="B4"><v>40</v></c><c r="C4"><v>400</v></c></row>
  </sheetData>${tableParts}
</worksheet>`
}

function columnCount(ref: string): number {
  const match = /^([A-Z]+)[0-9]+(?::([A-Z]+)[0-9]+)?$/.exec(ref)
  if (!match || !match[1]) return 1
  const toIndex = (letters: string): number =>
    [...letters].reduce((index, char) => index * 26 + char.charCodeAt(0) - 64, 0)
  return toIndex(match[2] ?? match[1]) - toIndex(match[1]) + 1
}

function tablePart(spec: TableSpec): string {
  const width = columnCount(spec.ref)
  const columns = Array.from(
    { length: width },
    (_, index) => `<tableColumn id="${index + 1}" name="Col${index + 1}"/>`,
  ).join('')
  const autoFilter = spec.autoFilter ?? `<autoFilter ref="${spec.ref}"/>`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" ` +
    `name="${spec.name}" displayName="${spec.name}" ref="${spec.ref}"${spec.attributes ?? ''}>` +
    `${autoFilter}<tableColumns count="${width}">${columns}</tableColumns>` +
    '<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>'
  )
}

async function fixture(tables: readonly TableSpec[] = defaultTables()): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', contentTypes(tables))
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets><sheet name="${SHEET}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  zip.file('xl/worksheets/sheet1.xml', worksheet(tables.length))
  if (tables.length > 0) {
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        tables
          .map(
            (table, index) =>
              `<Relationship Id="rId${index + 1}" Type="${TABLE_REL_TYPE}" Target="../tables/${table.path.split('/').pop()}"/>`,
          )
          .join('') +
        '</Relationships>',
    )
  }
  for (const table of tables) zip.file(table.path, tablePart(table))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function defaultTables(): TableSpec[] {
  return [{ path: 'xl/tables/table1.xml', name: 'Table1', ref: 'A1:C4' }]
}

type FilterRange = NonNullable<SheetFilterState['filter']>['range']

interface FilterOverrides {
  readonly range?: FilterRange
  readonly columns?: readonly FilterColumnState[]
  readonly filter?: SheetFilterState['filter']
  readonly hiddenRows?: readonly number[]
}

/// The whole-table filter snapshot the renderer emits for a table-origin
/// filter: the range covers the table ref exactly as the sidecar reads it
/// back at open time (totals band included when one exists).
function tableFilterState(overrides: FilterOverrides = {}): SheetFilterState {
  const range = overrides.range ?? { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 }
  return {
    sheetName: SHEET,
    filter:
      overrides.filter !== undefined
        ? overrides.filter
        : { range, columns: overrides.columns ?? [] },
    hiddenRows: [...(overrides.hiddenRows ?? [])],
    visibilityRange: range,
  }
}

/// planCellEditsToXlsx with everything between filterStates (slot 6) and
/// tableEdits (slot 27) left at its default.
async function planSave(
  source: EntrySource,
  filterStates: readonly SheetFilterState[],
  tableEdits: readonly {
    sheetName: string
    tableName: string
    name?: string
    remove?: true
  }[] = [],
) {
  return planCellEditsToXlsx(
    source,
    [],
    [],
    [],
    undefined,
    filterStates,
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    [],
    [],
    [],
    [],
    [],
    [],
    [],
    null,
    null,
    [],
    [],
    [],
    tableEdits,
  )
}

function autoFilterOccurrences(xml: string): number {
  return (xml.match(/<autoFilter\b/g) ?? []).length
}

/// Re-parses the saved table part the way the open path reads it back:
/// one entry per <filterColumn> with its colId and <filter> values.
function parseFilterColumns(xml: string): { colId: number; values: string[] }[] {
  return [...xml.matchAll(/<filterColumn colId="(\d+)">([\s\S]*?)<\/filterColumn>/g)].map(
    (match) => ({
      colId: Number(match[1]),
      values: [...(match[2] ?? '').matchAll(/<filter val="([^"]*)"/g)].map((value) => value[1]!),
    }),
  )
}

describe('table filter persistence', () => {
  it('routes a table-range filter into the table part and keeps the worksheet clean', async () => {
    const plan = await planSave(await createBufferEntrySource(await fixture()), [
      tableFilterState({ columns: [{ colId: 1, values: ['Beta'] }], hiddenRows: [2] }),
    ])
    const tableXml = plan.replaced.get('xl/tables/table1.xml')
    expect(tableXml).toBeDefined()
    expect(autoFilterOccurrences(tableXml!)).toBe(1)
    expect(tableXml).toContain(
      '<autoFilter ref="A1:C4"><filterColumn colId="1">' +
        '<filters><filter val="Beta"/></filters></filterColumn></autoFilter>',
    )
    const worksheetXml = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheetXml).toBeDefined()
    expect(worksheetXml).not.toContain('<autoFilter')
    expect(worksheetXml).toContain('<row r="3" hidden="1"')
  })

  it('round-trips: the saved part parses back into the snapshot and re-saves cleanly', async () => {
    const source = await fixture()
    const first = await planSave(await createBufferEntrySource(source), [
      tableFilterState({ columns: [{ colId: 1, values: ['Beta', 'Gamma'] }], hiddenRows: [3] }),
    ])
    const tableXml = first.replaced.get('xl/tables/table1.xml')!
    expect(parseFilterColumns(tableXml)).toEqual([{ colId: 1, values: ['Beta', 'Gamma'] }])
    const saved = (await assembleWithJsZip(source, first)).buffer

    // Re-open the saved bytes and change the filter: the existing element is
    // replaced, never duplicated.
    const second = await planSave(await createBufferEntrySource(saved), [
      tableFilterState({ columns: [{ colId: 0, values: ['Alpha'] }], hiddenRows: [1] }),
    ])
    const rewritten = second.replaced.get('xl/tables/table1.xml')!
    expect(autoFilterOccurrences(rewritten)).toBe(1)
    expect(parseFilterColumns(rewritten)).toEqual([{ colId: 0, values: ['Alpha'] }])
    expect(second.replaced.get('xl/worksheets/sheet1.xml')).not.toContain('<autoFilter')
  })

  it('stops the autoFilter ref above the totals band', async () => {
    const plan = await planSave(
      await createBufferEntrySource(
        await fixture([
          {
            path: 'xl/tables/table1.xml',
            name: 'Table1',
            ref: 'A1:C5',
            attributes: ' totalsRowCount="1" totalsRowShown="1"',
          },
        ]),
      ),
      [
        tableFilterState({
          range: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
          columns: [{ colId: 2, values: ['9'] }],
        }),
      ],
    )
    const tableXml = plan.replaced.get('xl/tables/table1.xml')!
    expect(tableXml).toContain('<autoFilter ref="A1:C4">')
    expect(tableXml).toContain('<filterColumn colId="2">')
  })

  it('keeps a bare autoFilter for a cleared or removed filter', async () => {
    const cleared = await planSave(await createBufferEntrySource(await fixture()), [
      tableFilterState({ columns: [], hiddenRows: [] }),
    ])
    expect(cleared.replaced.get('xl/tables/table1.xml')).toContain('<autoFilter ref="A1:C4"/>')

    const removed = await planSave(await createBufferEntrySource(await fixture()), [
      {
        sheetName: SHEET,
        filter: null,
        hiddenRows: [],
        visibilityRange: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
      },
    ])
    expect(removed.replaced.get('xl/tables/table1.xml')).toContain('<autoFilter ref="A1:C4"/>')
  })

  it('drops the autoFilter element for a headerless table', async () => {
    const plan = await planSave(
      await createBufferEntrySource(
        await fixture([
          {
            path: 'xl/tables/table1.xml',
            name: 'Table1',
            ref: 'A1:C4',
            attributes: ' headerRowCount="0"',
          },
        ]),
      ),
      [tableFilterState({ columns: [{ colId: 0, values: ['x'] }] })],
    )
    expect(plan.replaced.get('xl/tables/table1.xml')).not.toContain('<autoFilter')
  })

  it('still writes the worksheet autoFilter for a plain range without tables', async () => {
    const plan = await planSave(await createBufferEntrySource(await fixture([])), [
      tableFilterState({ columns: [{ colId: 1, values: ['Beta'] }], hiddenRows: [2] }),
    ])
    expect(plan.replaced.has('xl/tables/table1.xml')).toBe(false)
    const worksheetXml = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(worksheetXml).toContain('<autoFilter ref="A1:C4">')
    expect(worksheetXml).toContain('<filterColumn colId="1">')
    expect(worksheetXml).toContain('<row r="3" hidden="1"')
  })

  it('fails closed when the filter straddles two tables', async () => {
    const tables: TableSpec[] = [
      { path: 'xl/tables/table1.xml', name: 'Table1', ref: 'A1:B4' },
      { path: 'xl/tables/table2.xml', name: 'Table2', ref: 'D1:E4' },
    ]
    await expect(
      planSave(await createBufferEntrySource(await fixture(tables)), [
        tableFilterState({
          range: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 3 },
          columns: [{ colId: 0, values: ['x'] }],
        }),
      ]),
    ).rejects.toThrow(/overlaps Excel table/)
  })

  it('keeps routing after a same-save rename: the filter follows the table, not its name', async () => {
    const plan = await planSave(
      await createBufferEntrySource(await fixture()),
      [tableFilterState({ columns: [{ colId: 1, values: ['Beta'] }] })],
      [{ sheetName: SHEET, tableName: 'Table1', name: 'Renamed' }],
    )
    const tableXml = plan.replaced.get('xl/tables/table1.xml')!
    expect(tableXml).toContain('displayName="Renamed"')
    expect(parseFilterColumns(tableXml)).toEqual([{ colId: 1, values: ['Beta'] }])
    expect(plan.replaced.get('xl/worksheets/sheet1.xml')).not.toContain('<autoFilter')
  })

  it('falls back to the worksheet autoFilter after Convert to Range in the same save', async () => {
    const plan = await planSave(
      await createBufferEntrySource(await fixture()),
      [tableFilterState({ columns: [{ colId: 1, values: ['Beta'] }] })],
      [{ sheetName: SHEET, tableName: 'Table1', remove: true }],
    )
    expect(plan.removedEntries).toContain('xl/tables/table1.xml')
    const worksheetXml = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(worksheetXml).toContain('<autoFilter ref="A1:C4">')
    expect(worksheetXml).toContain('<filterColumn colId="1">')
  })
})
