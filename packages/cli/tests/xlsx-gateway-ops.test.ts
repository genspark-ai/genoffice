import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const sidecar = Boolean(xlsxSidecarPath())

// 1x1 red PNG
const PNG_1PX = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
)

async function parts(path: string): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  const out = new Map<string, string>()
  for (const name of Object.keys(zip.files)) {
    if (/\.(xml|rels|vml)$/.test(name)) out.set(name, await zip.file(name)!.async('string'))
  }
  return out
}

async function book(dir: string, rows: unknown[][] = DATA): Promise<string> {
  const table = join(dir, 'sales.json')
  writeFileSync(table, JSON.stringify(rows))
  const out = join(dir, 'sales.xlsx')
  expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
  return out
}

const DATA = [
  ['region', 'q1', 'q2'],
  ['North', 120, 150],
  ['South', 80, 95],
  ['West', 200, 210],
]

async function apply(path: string, ops: unknown[], extra: string[] = []) {
  const file = join(path, '..', `ops-${Math.random().toString(36).slice(2)}.json`)
  writeFileSync(file, JSON.stringify(ops))
  return run(['sheet', 'apply', path, '--ops', file, '--json', ...extra])
}

describe('genoffice sheet apply --ops: gateway-written features', () => {
  it('writes freeze panes, hidden rows/cols, page setup, hyperlinks and a note', async () => {
    const dir = tempDir()
    const out = await book(dir)
    const r = await apply(out, [
      { op: 'set_freeze', rows: 1, columns: 1 },
      { op: 'set_rows_hidden', row: 3, count: 1, hidden: true },
      { op: 'set_cols_hidden', column: 'C', hidden: true },
      { op: 'set_page_setup', orientation: 'landscape', scale: 80 },
      { op: 'set_page_setup', fitToWidth: 1, printGridlines: true },
      { op: 'set_hyperlink', address: 'A2', target: 'https://example.com/north' },
      { op: 'set_hyperlink', address: 'A3', target: 'sales!A1' },
      { op: 'set_note', address: 'B2', text: 'Provisional figure' },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail.features).toMatchObject({
      hyperlinks: 2,
      notes: 1,
      page_setup: 1,
      hidden_ranges: 2,
    })
    const p = await parts(out)
    const sheet = p.get('xl/worksheets/sheet1.xml')!
    expect(sheet).toMatch(/<pane [^>]*xSplit="1"[^>]*ySplit="1"[^>]*state="frozen"/)
    expect(sheet).toMatch(/<row r="3"[^>]*hidden="1"/)
    expect(sheet).toMatch(/<col [^>]*min="3"[^>]*max="3"[^>]*hidden="1"/)
    expect(sheet).toMatch(/<pageSetup [^>]*orientation="landscape"/)
    expect(sheet).toContain('<pageSetUpPr fitToPage="1"/>')
    expect(sheet).toMatch(/<pageSetup [^>]*fitToHeight="0"/)
    expect(sheet).not.toMatch(/<pageSetup [^>]*scale=/)
    expect(sheet).toMatch(/<printOptions [^>]*gridLines="1"/)
    expect(sheet).toMatch(/<hyperlink ref="A2" r:id="rId\d+"/)
    expect(sheet).toMatch(/<hyperlink ref="A3" location="sales!A1"/)
    expect(p.get('xl/worksheets/_rels/sheet1.xml.rels')).toContain('https://example.com/north')
    const comments = [...p.keys()].find((k) => /xl\/comments\d*\.xml$/.test(k))
    expect(comments).toBeDefined()
    expect(p.get(comments!)).toContain('Provisional figure')
    expect(p.get(comments!)).toContain('<author>GenOffice</author>')

    // a second batch keeps the first note (the whole comment set is rewritten from the file's state)
    const again = await apply(out, [{ op: 'set_note', address: 'C2', text: 'Second note' }])
    expect(again.code).toBe(0)
    const p2 = await parts(out)
    const c2 = p2.get([...p2.keys()].find((k) => /xl\/comments\d*\.xml$/.test(k))!)!
    expect(c2).toContain('Provisional figure')
    expect(c2).toContain('Second note')
  })

  it('creates a filter with criteria, then clears it', async () => {
    const dir = tempDir()
    const out = await book(dir)
    const r = await apply(out, [
      { op: 'set_filter', range: 'A1:C4' },
      { op: 'set_filter_criteria', column: 'A', values: ['North', 'West'] },
    ])
    expect(r.code).toBe(0)
    let sheet = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(sheet).toMatch(
      /<autoFilter ref="A1:C4"><filterColumn colId="0"><filters><filter val="North"\/><filter val="West"\/><\/filters>/,
    )
    expect(sheet).toMatch(/<row r="3"[^>]*hidden="1"/)
    expect(sheet).not.toMatch(/<row r="2"[^>]*hidden="1"/)

    const criteriaOnExisting = await apply(out, [
      { op: 'set_filter_criteria', column: 'B', values: ['120'] },
    ])
    expect(criteriaOnExisting.code).toBe(1)
    expect(criteriaOnExisting.json().message).toContain('already has criteria')

    // a fresh filter over the same range shows every row again, so new criteria can be set
    const refilter = await apply(out, [
      { op: 'set_filter', range: 'A1:C4' },
      { op: 'set_filter_criteria', column: 'B', values: ['80'] },
    ])
    expect(refilter.code).toBe(0)
    sheet = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(sheet).not.toMatch(/<row r="3"[^>]*hidden="1"/)
    expect(sheet).toMatch(/<row r="2"[^>]*hidden="1"/)
    expect(sheet).toMatch(/<row r="4"[^>]*hidden="1"/)

    const clearedThenCriteria = await apply(out, [
      { op: 'clear_filter' },
      { op: 'set_filter_criteria', column: 'A', values: ['North'] },
    ])
    expect(clearedThenCriteria.code).toBe(1)
    expect(clearedThenCriteria.json().message).toContain('cleared earlier in this batch')

    const cleared = await apply(out, [{ op: 'clear_filter' }])
    expect(cleared.code).toBe(0)
    sheet = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(sheet).not.toContain('<autoFilter')
    expect(sheet).not.toMatch(/hidden="1"/)
  })

  it('writes conditional formats and data validation, adding to the rules a sheet already has', async () => {
    const dir = tempDir()
    const out = await book(dir)
    const r = await apply(out, [
      {
        op: 'add_conditional_format',
        range: 'B2:C4',
        rule: {
          kind: 'number',
          operator: 'greaterThan',
          value: 100,
          format: { fillColor: '#C6EFCE', bold: true },
        },
      },
      {
        op: 'add_conditional_format',
        range: 'B2:B4',
        rule: { kind: 'colorScale', minColor: '#F8696B', maxColor: '#63BE7B' },
      },
      { op: 'add_conditional_format', range: 'C2:C4', rule: { kind: 'dataBar', color: '#638EC6' } },
      {
        op: 'add_conditional_format',
        range: 'A2:A4',
        rule: {
          kind: 'text',
          operator: 'contains',
          text: 'orth',
          format: { fontColor: '#9C0006' },
        },
      },
      {
        op: 'set_data_validation',
        range: 'A2:A4',
        validation: { kind: 'list', values: ['North', 'South', 'West'] },
      },
      {
        op: 'set_data_validation',
        range: 'B2:C4',
        validation: { kind: 'numberBetween', min: 0, max: 1000 },
      },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail.features).toMatchObject({ conditional_formats: 4, data_validations: 2 })
    const p = await parts(out)
    const sheet = p.get('xl/worksheets/sheet1.xml')!
    expect(sheet).toMatch(
      /<conditionalFormatting sqref="B2:C4"><cfRule type="cellIs" dxfId="\d+" priority="\d+" operator="greaterThan"><formula>100<\/formula>/,
    )
    expect(sheet).toMatch(/<conditionalFormatting sqref="B2:B4"><cfRule type="colorScale"/)
    expect(sheet).toMatch(/<conditionalFormatting sqref="C2:C4"><cfRule type="dataBar"/)
    expect(sheet).toMatch(/type="containsText"[^>]*text="orth"/)
    expect(p.get('xl/styles.xml')).toMatch(
      /<dxf><font><b\/><\/font><fill><patternFill><bgColor rgb="FFC6EFCE"\/>/,
    )
    expect(sheet).toMatch(
      /<dataValidation type="list"[^>]*sqref="A2:A4"><formula1>"North,South,West"<\/formula1>/,
    )
    expect(sheet).toMatch(
      /<dataValidation type="decimal"[^>]*sqref="B2:C4"><formula1>0<\/formula1><formula2>1000<\/formula2>/,
    )

    const more = await apply(out, [
      {
        op: 'add_conditional_format',
        range: 'A1:A1',
        rule: { kind: 'blank', blank: true, format: { fillColor: '#FFFF00' } },
      },
    ])
    expect(more.code).toBe(0)
    const appended = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(appended.match(/<conditionalFormatting\b/g)).toHaveLength(5)
    expect(appended).toMatch(/type="containsBlanks"[^>]*priority="5"/)

    const reset = await apply(out, [
      { op: 'clear_conditional_formats' },
      {
        op: 'add_conditional_format',
        range: 'A1:A1',
        rule: { kind: 'blank', blank: true, format: { fillColor: '#FFFF00' } },
      },
    ])
    expect(reset.code).toBe(0)
    const sheet2 = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(sheet2.match(/<conditionalFormatting\b/g)).toHaveLength(1)
    expect(sheet2).toContain('type="containsBlanks"')

    const dvAgain = await apply(out, [
      { op: 'set_data_validation', range: 'A5', validation: { kind: 'checkbox' } },
    ])
    expect(dvAgain.code).toBe(0)
    const sheet3 = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(sheet3).toContain('<dataValidations count="3">')
    expect(sheet3).toMatch(/<dataValidation type="decimal"[^>]*sqref="B2:C4">/)
    expect(sheet3).toMatch(/<dataValidation type="list"[^>]*sqref="A5">/)
    const fresh = await book(tempDir())
    const inBatch = await apply(fresh, [
      { op: 'set_data_validation', range: 'A5', validation: { kind: 'checkbox' } },
      { op: 'set_data_validation', range: 'A5', validation: null },
      { op: 'set_data_validation', range: 'A6', validation: null },
    ])
    expect(inBatch.code).toBe(0)
    const removals = (inBatch.json().warnings ?? [])
      .map((w: { message: string }) => w.message)
      .filter((m: string) => m.includes('nothing to remove'))
    expect(removals).toHaveLength(1)
    expect(removals[0]).toContain('A6')
    const dvDrop = await apply(out, [
      { op: 'set_data_validation', range: 'B2:C4', validation: null },
    ])
    expect(dvDrop.code).toBe(0)
    const sheet4 = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(sheet4).toContain('<dataValidations count="2">')
    expect(sheet4).not.toContain('sqref="B2:C4"')
  })

  it('adds a chart, a picture, a shape and an Excel table', async () => {
    const dir = tempDir()
    const out = await book(dir)
    const png = join(dir, 'logo.png')
    writeFileSync(png, PNG_1PX)
    const r = await apply(out, [
      {
        op: 'add_chart',
        chartType: 'column',
        dataRange: 'A1:C4',
        title: 'Sales by region',
        anchorCell: 'E2',
      },
      { op: 'add_image', path: png, anchorCell: 'E20' },
      { op: 'add_shape', shapeType: 'textbox', anchorCell: 'A8', text: 'Provisional' },
      { op: 'add_table', range: 'A1:C4', name: 'Sales', style: 'TableStyleMedium9' },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail.features).toMatchObject({ charts: 1, images: 1, shapes: 1, tables: 1 })
    const p = await parts(out)
    const chart = p.get('xl/charts/chart1.xml')!
    expect(chart).toContain('<c:barChart>')
    expect(chart).toContain('Sales by region')
    expect(chart).toMatch(/<c:f>'sales'!\$B\$2:\$B\$4<\/c:f>/)
    const drawing = p.get('xl/drawings/drawing1.xml')!
    expect(drawing).toContain('<xdr:graphicFrame')
    expect(drawing).toContain('<xdr:pic>')
    expect(drawing).toContain('Provisional')
    expect([...p.keys()].some((k) => k.startsWith('xl/media/'))).toBe(false)
    const zip = await JSZip.loadAsync(readFileSync(out))
    expect(Object.keys(zip.files).some((k) => /^xl\/media\/image\d+\.png$/.test(k))).toBe(true)
    const table = p.get('xl/tables/table1.xml')!
    expect(table).toMatch(/displayName="Sales"/)
    expect(table).toContain('ref="A1:C4"')
    expect(table).toContain('TableStyleMedium9')
    expect(table).toMatch(/<tableColumn id="1" name="region"\/>/)

    // default names are workbook-scoped: a table on another sheet skips names the file already has
    expect((await apply(out, [{ op: 'add_sheet', name: 'More' }])).code).toBe(0)
    expect(
      (
        await apply(out, [
          {
            op: 'set_range',
            sheet: 'More',
            range: 'A1:B3',
            values: [
              ['k', 'v'],
              ['a', 1],
              ['b', 2],
            ],
          },
        ])
      ).code,
    ).toBe(0)
    const tableWithWidth = await apply(out, [
      { op: 'set_col_width', sheet: 'More', column: 'A', widthPx: 120 },
      { op: 'add_table', sheet: 'More', range: 'A1:B3' },
    ])
    expect(tableWithWidth.code).toBe(1)
    expect(tableWithWidth.json().message).toContain('add_table cannot share a batch')
    expect((await apply(out, [{ op: 'add_table', sheet: 'More', range: 'A1:B3' }])).code).toBe(0)
    expect(
      (await apply(out, [{ op: 'add_table', sheet: 'sales', range: 'A1:C4', name: 'Sales' }])).code,
    ).toBe(1)
    const p2 = await parts(out)
    const names = [...p2.entries()]
      .filter(([k]) => k.startsWith('xl/tables/'))
      .map(([, v]) => /displayName="([^"]+)"/.exec(v)![1])
      .sort()
    expect(names).toEqual(['Sales', 'Table1'])
  })

  it('manages defined names, protection and sheet tabs', async () => {
    const dir = tempDir()
    const out = await book(dir)
    expect(
      (
        await apply(out, [
          { op: 'add_sheet', name: 'Notes' },
          { op: 'add_sheet', name: 'Archive' },
        ])
      ).code,
    ).toBe(0)
    // the gateway keeps name edits apart from sheet changes (sheet indexes move)
    expect(
      (await apply(out, [{ op: 'add_defined_name', name: 'Regions', ref: 'sales!$A$2:$A$4' }]))
        .code,
    ).toBe(0)
    const r = await apply(out, [
      { op: 'protect_sheet', sheet: 'Archive', protected: true },
      { op: 'set_sheet_hidden', sheet: 'Archive', hidden: true },
      { op: 'move_sheet', sheet: 'Notes', position: 1 },
      { op: 'duplicate_sheet', sheet: 'sales', name: 'sales copy' },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail.plan).toEqual([
      'Protect sheet',
      'Hide sheet Archive',
      'Move sheet Notes to position 1',
      'Duplicate sheet sales as "sales copy"',
    ])
    const p = await parts(out)
    const workbook = p.get('xl/workbook.xml')!
    expect(workbook).toContain('<definedName name="Regions">sales!$A$2:$A$4</definedName>')
    const names = [...workbook.matchAll(/<sheet [^>]*name="([^"]+)"/g)].map((m) => m[1])
    expect(names).toEqual(['Notes', 'sales', 'Archive', 'sales copy'])
    const archive = /<sheet [^>]*name="Archive"[^>]*\/>/.exec(workbook)![0]
    expect(archive).toContain('state="hidden"')
    const copy = [...p.entries()].find(
      ([k, v]) =>
        k.startsWith('xl/worksheets/sheet') &&
        v.includes('North') &&
        k !== 'xl/worksheets/sheet1.xml',
    )
    expect(copy).toBeDefined()
    expect([...p.values()].some((v) => v.includes('<sheetProtection'))).toBe(true)

    // a 31-character source still gets a distinct default copy name
    const longName = 'Quarterly regional sales 2026 v'
    expect(longName).toHaveLength(31)
    expect((await apply(out, [{ op: 'rename_sheet', sheet: 'Notes', name: longName }])).code).toBe(
      0,
    )
    const long = await apply(out, [{ op: 'duplicate_sheet', sheet: longName }])
    expect(long.code).toBe(0)
    expect((await parts(out)).get('xl/workbook.xml')).toContain(
      `name="${longName.slice(0, 27)} (2)"`,
    )

    // the copy is cloned from the file, so it cannot share a batch with cell edits
    const mixed = await apply(out, [
      { op: 'set_cell', sheet: 'sales', address: 'A9', value: 'x' },
      { op: 'duplicate_sheet', sheet: 'sales' },
    ])
    expect(mixed.code).toBe(1)
    expect(mixed.json().message).toContain('cannot share a batch with content ops')
    const namesWithSheets = await apply(out, [
      { op: 'add_defined_name', name: 'Later', ref: 'sales!$B$2' },
      { op: 'move_sheet', sheet: 'sales', position: 1 },
    ])
    expect(namesWithSheets.code).toBe(1)
    expect(namesWithSheets.json().message).toContain('defined-name ops cannot share a batch')
    const namesWithHide = await apply(out, [
      { op: 'add_defined_name', name: 'Later', ref: 'sales!$B$2' },
      { op: 'set_rows_hidden', sheet: 'sales', row: 2, hidden: true },
    ])
    expect(namesWithHide.code).toBe(1)
    expect(namesWithHide.json().message).toContain('defined-name ops cannot share a batch')

    const gone = await apply(out, [{ op: 'delete_defined_name', name: 'Regions' }])
    expect(gone.code).toBe(0)
    expect((await parts(out)).get('xl/workbook.xml')).not.toContain('Regions')
    const missing = await apply(out, [{ op: 'delete_defined_name', name: 'Nope' }])
    expect(missing.code).toBe(1)
  })

  it('refuses editor-only ops with the reason and keeps edit_chart to file charts', async () => {
    const dir = tempDir()
    const out = await book(dir)
    const pivot = await apply(out, [{ op: 'refresh_pivot' }])
    expect(pivot.code).toBe(1)
    expect(pivot.json().message).toContain('computed by the editor')
    expect(pivot.json().detail.not_available).toContain('refresh_pivot')
    const demo = await apply(out, [{ op: 'edit_chart', chartPath: 'added-chart-1', title: 'x' }])
    expect(demo.code).toBe(1)
    expect(demo.json().message).toContain('xl/charts/chartN.xml')

    expect(
      (await apply(out, [{ op: 'add_chart', chartType: 'pie', dataRange: 'A1:B4' }])).code,
    ).toBe(0)
    const retitled = await apply(out, [
      {
        op: 'edit_chart',
        chartPath: 'xl/charts/chart1.xml',
        title: 'Share by region',
        legend: 'bottom',
      },
    ])
    expect(retitled.code).toBe(0)
    const chart = (await parts(out)).get('xl/charts/chart1.xml')!
    expect(chart).toContain('Share by region')
    expect(chart).toContain('<c:legendPos val="b"/>')
  })
})

describe('genoffice sheet read: features and --formats', () => {
  it.skipIf(!xlsxSidecarPath())(
    'reports panes, filter, merges, charts and cell formats',
    async () => {
      const dir = tempDir()
      const out = await book(dir)
      expect(
        (
          await apply(out, [
            { op: 'format_range', range: 'A1:C1', format: { bold: true, fillColor: '#FFFF00' } },
            { op: 'format_range', range: 'B2:C4', format: { numberFormat: '#,##0.00' } },
            { op: 'merge_cells', range: 'B4:C4' },
            { op: 'set_col_width', column: 'A', widthPx: 140 },
            { op: 'set_row_height', row: 1, heightPoints: 30 },
            { op: 'set_freeze', rows: 1, columns: 0 },
            { op: 'set_filter', range: 'A1:C4' },
            { op: 'add_chart', chartType: 'line', dataRange: 'A1:C4', title: 'Trend' },
            { op: 'set_hyperlink', address: 'A2', target: 'https://example.com' },
            { op: 'set_note', address: 'A3', text: 'note' },
          ])
        ).code,
      ).toBe(0)
      const plain = await run(['sheet', 'read', out, '--json'])
      expect(plain.code).toBe(0)
      expect(plain.json().detail.features).toMatchObject({
        frozen: { rows: 1, columns: 0 },
        filter: 'A1:C4',
        merges: ['B4:C4'],
        charts: [{ title: 'Trend', types: ['lineChart'], series: 2 }],
        hyperlinks: 1,
        notes: 1,
      })
      expect(plain.json().detail.formats).toBeUndefined()

      // a range outside the used area still reports the sheet's features
      const beyond = await run(['sheet', 'read', out, '--range', 'H40:J42', '--json'])
      expect(beyond.code).toBe(0)
      expect(beyond.json().detail.rows).toEqual([
        [null, null, null],
        [null, null, null],
        [null, null, null],
      ])
      expect(beyond.json().detail.features).toMatchObject({
        filter: 'A1:C4',
        frozen: { rows: 1, columns: 0 },
      })

      const rich = await run(['sheet', 'read', out, '--formats', '--json'])
      expect(rich.code).toBe(0)
      const d = rich.json().detail
      expect(d.formats.A1).toMatchObject({ bold: true, fillColor: '#FFFF00' })
      expect(d.formats.B2).toMatchObject({ numberFormat: '#,##0.00' })
      expect(d.formats.A2).toBeUndefined()
      expect(d.columnWidths.A).toBe(140)
      expect(d.rowHeights['1']).toBe(30)
    },
  )
})

describe('genoffice sheet apply --ops: pivots, sparklines, print setup', () => {
  const SALES = [
    ['region', 'product', 'amount'],
    ['North', 'A', 10],
    ['South', 'B', 20],
    ['North', 'B', 5],
    ['West', 'A', 7],
  ]

  it('adds a pivot table: baked cells plus the native definition, cache and records', async () => {
    const dir = tempDir()
    const out = await book(dir, SALES)
    const r = await apply(out, [
      {
        op: 'add_pivot',
        sourceRange: 'A1:C5',
        targetCell: 'E1',
        rowFields: 'region',
        values: [{ field: 'amount', agg: 'sum', numFmt: '#,##0.00' }],
        name: 'ByRegion',
      },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail.features).toMatchObject({ pivots: 1 })
    const p = await parts(out)
    const definition = p.get('xl/pivotTables/pivotTable1.xml')!
    expect(definition).toContain('name="ByRegion"')
    expect(definition).toContain('<location ref="E1:F5"')
    expect(p.get('xl/pivotCache/pivotCacheDefinition1.xml')).toContain('sheet="sales"')
    expect(p.get('xl/pivotCache/pivotCacheRecords1.xml')).toContain('count="4"')
    expect(p.get('xl/workbook.xml')).toContain('<pivotCaches>')
    expect(p.get('[Content_Types].xml')).toContain('pivotTable+xml')
    const sheet = p.get('xl/worksheets/sheet1.xml')!
    expect(sheet).toMatch(/<c r="F2"[^>]*>\s*<v>15<\/v>/)
    expect(sheet).toMatch(/<c r="F5"[^>]*>\s*<v>42<\/v>/)
    if (sidecar) {
      const read = await run(['sheet', 'read', out, '--range', 'E1:F5', '--json'])
      expect(read.json().detail.rows).toEqual([
        ['region', 'Sum of amount'],
        ['North', 15],
        ['South', 20],
        ['West', 7],
        ['Grand Total', 42],
      ])
      expect(read.json().detail.features.pivots).toEqual(['E1:F5'])
    }

    const overlap = await apply(out, [
      {
        op: 'add_pivot',
        sourceRange: 'A1:C5',
        targetCell: 'F3',
        rowFields: 'product',
        values: [{ field: 'amount', agg: 'count' }],
      },
    ])
    expect(overlap.code).toBe(1)
    expect(overlap.json().message).toContain('overlap pivot "ByRegion"')
    const occupied = await apply(out, [
      {
        op: 'add_pivot',
        sourceRange: 'A1:C5',
        targetCell: 'B3',
        rowFields: 'product',
        values: [{ field: 'amount', agg: 'count' }],
      },
    ])
    expect(occupied.code).toBe(1)
    expect(occupied.json().message).toContain('overlap its source')
    const badField = await apply(out, [
      {
        op: 'add_pivot',
        sourceRange: 'A1:C5',
        targetCell: 'H1',
        rowFields: 'nope',
        values: [{ field: 'amount', agg: 'sum' }],
      },
    ])
    expect(badField.code).toBe(1)
    expect(badField.json().message).toContain('"nope" is not a source header')
  })

  it('places a pivot with a column field on another sheet and refuses a sheet born in the batch', async () => {
    const dir = tempDir()
    const out = await book(dir, SALES)
    expect((await apply(out, [{ op: 'add_sheet', name: 'Summary' }])).code).toBe(0)
    const r = await apply(out, [
      {
        op: 'add_pivot',
        sourceRange: 'A1:C5',
        targetSheet: 'Summary',
        targetCell: 'A1',
        rowFields: 'region',
        columnField: 'product',
        values: [{ field: 'amount', agg: 'sum' }],
      },
    ])
    expect(r.code).toBe(0)
    const p = await parts(out)
    expect(p.get('xl/pivotTables/pivotTable1.xml')).toContain('<colFields count="1">')
    const summary = p.get('xl/worksheets/sheet2.xml')!
    expect(summary).toMatch(/<c r="B2"[^>]*>\s*<v>10<\/v>/)
    expect(summary).toMatch(/<c r="D2"[^>]*>\s*<v>15<\/v>/)
    if (sidecar) {
      const read = await run(['sheet', 'read', out, '--sheet', 'Summary', '--json'])
      expect(read.json().detail.rows[0]).toEqual(['region', 'A', 'B', 'Grand Total'])
      expect(read.json().detail.rows[1]).toEqual(['North', 10, 5, 15])
    }

    const born = await apply(out, [
      { op: 'add_sheet', name: 'Later' },
      {
        op: 'add_pivot',
        sourceRange: 'A1:C5',
        targetSheet: 'Later',
        targetCell: 'A1',
        rowFields: 'region',
        values: [{ field: 'amount', agg: 'sum' }],
      },
    ])
    expect(born.code).toBe(1)
    expect(born.json().message).toMatch(/sheet not found: Later|added in this batch/)
  })

  it('adds sparklines and print setup: titles, header/footer and manual breaks', async () => {
    const dir = tempDir()
    const out = await book(dir, [
      ['name', 'q1', 'q2', 'q3'],
      ['a', 1, 2, 3],
      ['b', 3, 2, 1],
    ])
    const r = await apply(out, [
      { op: 'add_sparkline', dataRange: 'B2:D3', type: 'column', color: '#FF0000' },
      {
        op: 'set_page_setup',
        printTitles: '1:1',
        header: { center: 'Quarterly' },
        footer: { right: 'Page &P of &N' },
        rowBreaks: [2],
      },
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail.features).toMatchObject({ sparklines: 2, page_setup: 1 })
    const p = await parts(out)
    const sheet = p.get('xl/worksheets/sheet1.xml')!
    expect(sheet).toContain('<x14:sparklineGroup displayEmptyCellsAs="gap" type="column">')
    expect(sheet).toContain("<xm:f>'sales'!$B$2:$D$2</xm:f><xm:sqref>E2</xm:sqref>")
    expect(sheet).toContain('<xm:sqref>E3</xm:sqref>')
    expect(sheet).toContain('<x14:colorSeries rgb="FFFF0000"/>')
    expect(sheet).toContain('<oddHeader>&amp;CQuarterly</oddHeader>')
    expect(sheet).toContain('<oddFooter>&amp;RPage &amp;P of &amp;N</oddFooter>')
    expect(sheet).toContain(
      '<rowBreaks count="1" manualBreakCount="1"><brk id="2" max="16383" man="1"/>',
    )
    expect(p.get('xl/workbook.xml')).toMatch(/_xlnm\.Print_Titles"[^>]*>'?sales'?!\$1:\$1</)

    const twice = await apply(out, [{ op: 'add_sparkline', dataRange: 'B2:D2', type: 'line' }])
    expect(twice.code).toBe(1)
    expect(twice.json().message).toContain('already has a sparkline')
  })

  it('refuses a sparkline range past 200 rows instead of dropping the rest', async () => {
    const dir = tempDir()
    const rows = Array.from({ length: 201 }, (_, i) => [`r${i}`, i, i + 1])
    const out = await book(dir, [['name', 'q1', 'q2'], ...rows])
    const r = await apply(out, [{ op: 'add_sparkline', dataRange: 'B2:C202', type: 'line' }])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('spans 201 rows; one op writes at most 200 sparklines')
    const ok = await apply(out, [{ op: 'add_sparkline', dataRange: 'B2:C201', type: 'line' }])
    expect(ok.code).toBe(0)
    expect(ok.json().detail.features).toMatchObject({ sparklines: 200 })
    expect((await parts(out)).get('xl/worksheets/sheet1.xml')).toContain(
      '<xm:sqref>D201</xm:sqref>',
    )
  })

  it('refuses a pivot target holding formulas and source formulas set in the same batch', async () => {
    const dir = tempDir()
    const out = await book(dir, SALES)
    expect((await apply(out, [{ op: 'set_formula', address: 'F3', formula: '=C2*2' }])).code).toBe(
      0,
    )
    const pivot = {
      op: 'add_pivot',
      sourceRange: 'A1:C5',
      targetCell: 'E1',
      rowFields: 'region',
      values: [{ field: 'amount', agg: 'sum' }],
    }
    const onFormula = await apply(out, [pivot])
    expect(onFormula.code).toBe(1)
    expect(onFormula.json().message).toContain('would overwrite 1 non-empty cell(s) (F3)')
    const sameBatch = await apply(out, [
      { op: 'set_formula', address: 'C6', formula: '=C2+C3' },
      { ...pivot, sourceRange: 'A1:C6', targetCell: 'H1' },
    ])
    expect(sameBatch.code).toBe(1)
    expect(sameBatch.json().message).toContain(
      'the source has formulas (C6) and this batch edits cells (sales!C6)',
    )
    // a precedent edited in the batch: the engine would evaluate the formula against the old input
    expect((await apply(out, [{ op: 'set_formula', address: 'C5', formula: '=B5*0' }])).code).toBe(
      0,
    )
    const precedent = await apply(out, [
      { op: 'set_cell', address: 'B5', value: 3 },
      { ...pivot, targetCell: 'H1' },
    ])
    expect(precedent.code).toBe(1)
    expect(precedent.json().message).toContain(
      'the source has formulas (C5) and this batch edits cells (sales!B5)',
    )
  })

  it('convert_to_values keeps later ops pinned to their own index', async () => {
    const dir = tempDir()
    const out = await book(dir, SALES)
    // no formulas in the range: the op is dropped from the batch, the rest must not shift
    const r = await apply(
      out,
      [
        { op: 'convert_to_values', range: 'C2:C5' },
        { op: 'set_cell', address: 'E1', value: 'ok' },
        { op: 'set_cell', sheet: 'Nope', address: 'A1', value: 'x' },
      ],
      ['--best-effort'],
    )
    expect(r.code).toBe(0)
    expect(r.json().status).toBe('partial')
    expect(r.json().detail.failures).toEqual([
      expect.objectContaining({ index: 2, op: 'set_cell' }),
    ])
    expect(r.json().detail.batch).toMatchObject({ total: 3, applied: 2, failed: 1 })
    // a malformed convert_to_values is pinned too, so best effort can drop just that op
    const bad = await apply(
      out,
      [
        { op: 'set_cell', address: 'E2', value: 'still ok' },
        { op: 'convert_to_values', range: 'not-a-range' },
      ],
      ['--best-effort'],
    )
    expect(bad.code).toBe(0)
    expect(bad.json().detail.failures).toEqual([
      expect.objectContaining({ index: 1, op: 'convert_to_values' }),
    ])
  })

  it.skipIf(!sidecar)('aggregates source formulas by their computed results', async () => {
    const dir = tempDir()
    const out = await book(dir, [
      ['region', 'qty', 'price', 'amount'],
      ['North', 2, 5, '=B2*C2'],
      ['South', 4, 5, '=B3*C3'],
      ['North', 1, 5, '=B4*C4'],
    ])
    const pivot = {
      op: 'add_pivot',
      sourceRange: 'A1:D4',
      targetCell: 'F1',
      rowFields: 'region',
      values: [{ field: 'amount', agg: 'sum' }],
    }
    const r = await apply(out, [pivot])
    expect(r.code).toBe(0)
    const sheet = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(sheet).toMatch(/<c r="G2"[^>]*>\s*<v>15<\/v>/)
    expect(sheet).toMatch(/<c r="G3"[^>]*>\s*<v>20<\/v>/)
    expect(sheet).toMatch(/<c r="G4"[^>]*>\s*<v>35<\/v>/)
    expect((await parts(out)).get('xl/pivotCache/pivotCacheRecords1.xml')).toContain('<n v="10"/>')

    expect((await apply(out, [{ op: 'set_formula', address: 'D3', formula: '=B3/0' }])).code).toBe(
      0,
    )
    const broken = await apply(out, [{ ...pivot, targetCell: 'J1' }])
    expect(broken.code).toBe(1)
    expect(broken.json().message).toContain('source formula(s) D3 evaluate to an error')
  })

  it.skipIf(!sidecar)(
    'convert_to_values freezes formula results and keeps error formulas',
    async () => {
      const dir = tempDir()
      const out = await book(dir, [
        ['a', 'b', 'total'],
        [2, 3, '=A2*B2'],
        [4, 0, '=A3/B3'],
      ])
      const r = await apply(out, [{ op: 'convert_to_values', range: 'C2:C3' }])
      expect(r.code).toBe(0)
      expect(r.json().warnings?.[0]?.message).toContain('evaluate to an error')
      const sheet = (await parts(out)).get('xl/worksheets/sheet1.xml')!
      expect(sheet).toContain('<c r="C2"')
      expect(sheet).not.toMatch(/<c r="C2"[^>]*>\s*<f>/)
      expect(sheet).toMatch(/<c r="C3"[^>]*>\s*<f>A3\/B3<\/f>/)
      const read = await run(['sheet', 'read', out, '--range', 'C2:C2', '--json'])
      expect(read.json().detail.rows).toEqual([[6]])
      expect(read.json().detail.formulas ?? {}).toEqual({})
      // edits after the convert are fine: the frozen value is the one the convert saw
      const later = await book(tempDir(), [
        ['a', 'b', 'total'],
        [2, 3, '=A2*B2'],
      ])
      const r2 = await apply(later, [
        { op: 'convert_to_values', range: 'C2:C2' },
        { op: 'set_cell', address: 'A2', value: 10 },
      ])
      expect(r2.code).toBe(0)
      const read2 = await run(['sheet', 'read', later, '--range', 'A2:C2', '--json'])
      expect(read2.json().detail.rows).toEqual([[10, 3, 6]])
      // an earlier convert's own writes are the engine's results, not edits a later convert could miss
      const twice = await book(tempDir(), [
        ['a', 'b', 'total', 'double'],
        [2, 3, '=A2*B2', '=C2*2'],
      ])
      const r3 = await apply(twice, [
        { op: 'convert_to_values', range: 'C2:C2' },
        { op: 'convert_to_values', range: 'D2:D2' },
      ])
      expect(r3.code).toBe(0)
      const read3 = await run(['sheet', 'read', twice, '--range', 'C2:D2', '--json'])
      expect(read3.json().detail.rows).toEqual([[6, 12]])
      expect(read3.json().detail.formulas ?? {}).toEqual({})
      // but a frozen cell rewritten before the next convert is an edit again
      const rewritten = await book(tempDir(), [
        ['a', 'b', 'total', 'double'],
        [2, 3, '=A2*B2', '=C2*2'],
      ])
      const r4 = await apply(rewritten, [
        { op: 'convert_to_values', range: 'C2:C2' },
        { op: 'set_cell', address: 'C2', value: 100 },
        { op: 'convert_to_values', range: 'D2:D2' },
      ])
      expect(r4.code).toBe(1)
      expect(r4.json().message).toContain(
        'ops[2] (convert_to_values): earlier ops in this batch edit cells (sales!C2)',
      )
    },
  )

  it('convert_to_values refuses cells edited earlier in the same batch', async () => {
    const dir = tempDir()
    const out = await book(dir, [
      ['a', 'b', 'total'],
      [2, 3, '=A2*B2'],
    ])
    // the engine evaluates from the file: a precedent edited in the batch would freeze a stale result
    const precedent = await apply(out, [
      { op: 'set_cell', address: 'A2', value: 10 },
      { op: 'convert_to_values', range: 'C2:C2' },
    ])
    expect(precedent.code).toBe(1)
    expect(precedent.json().message).toContain(
      'ops[1] (convert_to_values): earlier ops in this batch edit cells (sales!A2)',
    )
    // a formula written in the batch is invisible to the engine too
    const written = await apply(out, [
      { op: 'set_formula', address: 'C3', formula: '=A2+B2' },
      { op: 'convert_to_values', range: 'C2:C3' },
    ])
    expect(written.code).toBe(1)
    expect(written.json().message).toContain('edit cells (sales!C3)')
    const sheet = (await parts(out)).get('xl/worksheets/sheet1.xml')!
    expect(sheet).toMatch(/<c r="C2"[^>]*>\s*<f>A2\*B2<\/f>/)
  })
})
