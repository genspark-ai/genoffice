import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, assertOnlyTouchedEntriesChanged } from '../src/gateway/xlsx-gateway'
import { applyFilterState, FilterEditError } from '../src/gateway/xlsx-filter'
import { buildStructureFixture } from './fixture-builder'

const AREA = { startRow: 0, endRow: 9, startColumn: 0, endColumn: 3 }

const worksheet = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>
    <row r="2"><c r="A2"><v>1</v></c></row>
    <row r="3" hidden="1"><c r="A3"><v>2</v></c></row>
    <row r="5" ht="20" customHeight="1"><c r="A5"><v>4</v></c></row>
  </sheetData>
</worksheet>`

describe('applyFilterState', () => {
  it('writes a values filter with hidden rows and creates missing row elements', () => {
    const xml = applyFilterState(worksheet, {
      sheetName: 'Data',
      filter: {
        range: AREA,
        columns: [{ colId: 0, values: ['1', '4'], blank: true }],
      },
      hiddenRows: [2, 3],
      visibilityRange: AREA,
    })
    expect(xml).toContain(
      '<autoFilter ref="A1:D10"><filterColumn colId="0">' +
        '<filters blank="1"><filter val="1"/><filter val="4"/></filters>' +
        '</filterColumn></autoFilter>',
    )
    // Row 3 (index 2) stays hidden, row 4 (index 3) has no element yet and is
    // created hidden, row 5 keeps its height and gets unhidden semantics.
    expect(xml).toContain('<row r="3" hidden="1">')
    expect(xml).toContain('<row r="4" hidden="1"/>')
    expect(xml).toContain('<row r="5" ht="20" customHeight="1">')
    // Header row is never touched.
    expect(xml).toContain('<row r="1">')
  })

  it('unhides previously hidden rows that the filter no longer excludes', () => {
    const xml = applyFilterState(worksheet, {
      sheetName: 'Data',
      filter: { range: AREA, columns: [] },
      hiddenRows: [],
      visibilityRange: AREA,
    })
    expect(xml).toContain('<row r="3"><c r="A3">')
    expect(xml).toContain('<autoFilter ref="A1:D10"/>')
  })

  it('replaces an existing autoFilter in place and can remove it', () => {
    const withFilter = `<worksheet><sheetData><row r="2" hidden="1"><c r="A2"><v>1</v></c></row></sheetData><autoFilter ref="A1:B5"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter><mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells></worksheet>`
    const replaced = applyFilterState(withFilter, {
      sheetName: 'Data',
      filter: { range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 }, columns: [] },
      hiddenRows: [],
      visibilityRange: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
    })
    expect(replaced).toContain('<autoFilter ref="A1:B5"/><mergeCells')
    expect(replaced).not.toContain('filterColumn')
    expect(replaced).toContain('<row r="2"><c r="A2">')

    const removed = applyFilterState(withFilter, {
      sheetName: 'Data',
      filter: null,
      hiddenRows: [],
      visibilityRange: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
    })
    expect(removed).not.toContain('autoFilter')
    expect(removed).toContain('<row r="2"><c r="A2">')
  })

  it('serializes custom criteria and rejects unknown operators', () => {
    const xml = applyFilterState(worksheet, {
      sheetName: 'Data',
      filter: {
        range: AREA,
        columns: [
          {
            colId: 2,
            customs: {
              and: true,
              filters: [
                { val: 5, operator: 'greaterThan' },
                { val: '*end', operator: 'equal' },
              ],
            },
          },
        ],
      },
      hiddenRows: [],
      visibilityRange: AREA,
    })
    expect(xml).toContain(
      '<filterColumn colId="2"><customFilters and="1">' +
        '<customFilter operator="greaterThan" val="5"/><customFilter val="*end"/>' +
        '</customFilters></filterColumn>',
    )
    expect(() =>
      applyFilterState(worksheet, {
        sheetName: 'Data',
        filter: {
          range: AREA,
          columns: [{ colId: 0, customs: { filters: [{ val: 1, operator: 'aboveAverage' }] } }],
        },
        hiddenRows: [],
        visibilityRange: AREA,
      }),
    ).toThrow(FilterEditError)
  })

  it('escapes filter values', () => {
    const xml = applyFilterState(worksheet, {
      sheetName: 'Data',
      filter: {
        range: AREA,
        columns: [{ colId: 0, values: ['a<b>&"\''] }],
      },
      hiddenRows: [],
      visibilityRange: AREA,
    })
    expect(xml).toContain('<filter val="a&lt;b&gt;&amp;&quot;&apos;"/>')
  })
})

describe('filter save integration', () => {
  it('writes the autoFilter through the preservation pipeline', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildStructureFixture(),
      [],
      [],
      [],
      undefined,
      [
        {
          sheetName: 'Data',
          filter: {
            range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 3 },
            columns: [{ colId: 0, values: ['1', '10'] }],
          },
          hiddenRows: [1, 3],
          visibilityRange: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 3 },
        },
      ],
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect(mutation.touchedEntries).toContain('xl/worksheets/sheet1.xml')
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<autoFilter ref="A1:D10"><filterColumn colId="0">')
    expect(sheet).toContain('<row r="2" hidden="1">')
    expect(sheet).toContain('<row r="4" hidden="1">')
    expect(sheet).toContain('<row r="10"><c r="A10">')
  })
})
