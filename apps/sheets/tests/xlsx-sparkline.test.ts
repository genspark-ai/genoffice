import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  assembleWithJsZip,
  assertOnlyTouchedEntriesChanged,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '../src/gateway/xlsx-gateway'
import type { SheetSparklineAddition } from '../src/gateway/xlsx-gateway'
import { applySparklineAdditions, SparklineAddError } from '../src/gateway/xlsx-sparkline'
import { buildEditFixture } from './fixture-builder'

const X14_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'

const bareWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`

function group(overrides: Partial<Omit<SheetSparklineAddition, 'sheetName'>> = {}) {
  return {
    type: 'line' as const,
    cells: [{ cell: 'D2', sourceRef: 'Sheet1!A2:F2' }],
    ...overrides,
  }
}

describe('applySparklineAdditions', () => {
  it('creates extLst, the x14 ext, and a full group', () => {
    const xml = applySparklineAdditions(bareWorksheet, [
      group({
        type: 'column',
        cells: [
          { cell: 'D2', sourceRef: 'Sheet1!A2:F2' },
          { cell: 'D3', sourceRef: 'Sheet1!A3:F3' },
        ],
      }),
    ])
    expect(xml).toContain(
      `<extLst><ext uri="${X14_URI}"` +
        ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">',
    )
    expect(xml).toContain(
      '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">',
    )
    expect(xml).toContain('<x14:sparklineGroup displayEmptyCellsAs="gap" type="column">')
    expect(xml).toContain('<x14:colorSeries rgb="FF376092"/>')
    expect(xml).toContain('<x14:colorNegative rgb="FFD00000"/>')
    expect(xml).toContain(
      '<x14:sparkline><xm:f>Sheet1!A2:F2</xm:f><xm:sqref>D2</xm:sqref></x14:sparkline>',
    )
    expect(xml).toContain(
      '<x14:sparkline><xm:f>Sheet1!A3:F3</xm:f><xm:sqref>D3</xm:sqref></x14:sparkline>',
    )
    expect(xml).toMatch(/<\/extLst><\/worksheet>$/)
  })

  it('omits the type attribute for line groups and maps colors to ARGB', () => {
    const xml = applySparklineAdditions(bareWorksheet, [group({ color: '#10b981' })])
    expect(xml).toContain('<x14:sparklineGroup displayEmptyCellsAs="gap">')
    expect(xml).not.toContain('type="line"')
    expect(xml).toContain('<x14:colorSeries rgb="FF10B981"/>')
  })

  it('writes stacked groups with their type and escapes source refs', () => {
    const xml = applySparklineAdditions(bareWorksheet, [
      group({
        type: 'stacked',
        cells: [{ cell: 'B1', sourceRef: "'P&L'!A1:F1" }],
      }),
    ])
    expect(xml).toContain('type="stacked"')
    expect(xml).toContain("<xm:f>'P&amp;L'!A1:F1</xm:f>")
  })

  it('appends the ext to an existing extLst with foreign exts intact', () => {
    const existing = bareWorksheet.replace(
      '</worksheet>',
      '<extLst><ext uri="{OTHER}"><foo/></ext></extLst></worksheet>',
    )
    const xml = applySparklineAdditions(existing, [group()])
    expect(xml).toContain('<ext uri="{OTHER}"><foo/></ext>')
    expect(xml).toMatch(/<foo\/><\/ext><ext uri="\{05C60535/)
    expect((xml.match(/<extLst>/g) ?? []).length).toBe(1)
  })

  it('appends new groups inside an existing x14 sparkline ext', () => {
    const existing = bareWorksheet.replace(
      '</worksheet>',
      `<extLst><ext uri="${X14_URI}"` +
        ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
        '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
        '<x14:sparklineGroup displayEmptyCellsAs="gap"><x14:sparklines>' +
        '<x14:sparkline><xm:f>Sheet1!A1:F1</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline>' +
        '</x14:sparklines></x14:sparklineGroup>' +
        '</x14:sparklineGroups></ext></extLst></worksheet>',
    )
    const xml = applySparklineAdditions(existing, [group()])
    expect((xml.match(/<ext /g) ?? []).length).toBe(1)
    expect((xml.match(/<x14:sparklineGroups /g) ?? []).length).toBe(1)
    expect((xml.match(/<x14:sparklineGroup /g) ?? []).length).toBe(2)
    expect(xml).toContain('<xm:sqref>C1</xm:sqref>')
    expect(xml).toContain('<xm:sqref>D2</xm:sqref>')
  })

  it('fails closed when the host cell already has a sparkline', () => {
    const existing = applySparklineAdditions(bareWorksheet, [group()])
    expect(() => applySparklineAdditions(existing, [group()])).toThrow(SparklineAddError)
    expect(() => applySparklineAdditions(existing, [group()])).toThrow(/D2 already has a sparkline/)
    // D20 must not collide with D2.
    expect(() =>
      applySparklineAdditions(existing, [
        group({
          cells: [{ cell: 'D20', sourceRef: 'Sheet1!A20:F20' }],
        }),
      ]),
    ).not.toThrow()
  })

  it('rejects duplicate host cells within one save', () => {
    expect(() =>
      applySparklineAdditions(bareWorksheet, [group(), group({ type: 'column' })]),
    ).toThrow(/D2 already has a sparkline/)
  })
})

describe('sparkline save integration', () => {
  it('writes the ext through the preservation pipeline', async () => {
    const buffer = await buildEditFixture()
    const source = await createBufferEntrySource(buffer)
    const plan = await planCellEditsToXlsx(
      source,
      [],
      [],
      [],
      undefined,
      [],
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
      [
        { sheetName: 'Data', ...group() },
        {
          sheetName: 'Data',
          ...group({ type: 'column', cells: [{ cell: 'E5', sourceRef: 'Data!A5:C5' }] }),
        },
      ],
    )
    const mutation = await assembleWithJsZip(buffer, plan)
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect(mutation.touchedEntries).toContain('xl/worksheets/sheet1.xml')

    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain(`<ext uri="${X14_URI}"`)
    // Both request groups land in one sparklineGroups container.
    expect((sheet?.match(/<x14:sparklineGroups /g) ?? []).length).toBe(1)
    expect((sheet?.match(/<x14:sparklineGroup /g) ?? []).length).toBe(2)
    expect(sheet).toContain('<xm:sqref>D2</xm:sqref>')
    expect(sheet).toContain('<xm:sqref>E5</xm:sqref>')
  })
})
