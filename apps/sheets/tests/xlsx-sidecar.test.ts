import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { z } from 'zod'

import { blankXlsxBuffer } from '../src/gateway/csv-import'
import { saveWorkbookViaSidecar } from '../src/gateway/xlsx-package-io'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { workbookRangeResultSchema } from '../src/shared/desktop-api'
import { buildCompatibilityFixture } from './fixture-builder'

const openResultSchema = z.object({
  sessionId: z.string().uuid(),
  name: z.string(),
  entryCount: z.number().int().nonnegative(),
  sheets: z.array(
    z.object({
      id: z.string(),
      name: z.string(),
      rowCount: z.number().int().positive(),
      columnCount: z.number().int().positive(),
    }),
  ),
})

describe('XLSX Rust sidecar', () => {
  it('opens a workbook and reads a sparse cell range', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-sidecar-test-'))
    const path = join(directory, 'fixture.xlsx')
    await writeFile(path, await buildCompatibilityFixture())
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    let sessionId: string | null = null
    try {
      const opened = openResultSchema.parse(await client.open(path))
      sessionId = opened.sessionId
      expect(opened.sheets).toEqual([
        {
          id: 'sheet-1',
          name: 'Sheet1',
          rowCount: 1,
          columnCount: 2,
        },
      ])

      const result = workbookRangeResultSchema.parse(
        await client.readRange({
          sessionId,
          sheetId: 'sheet-1',
          range: {
            startRow: 0,
            endRow: 0,
            startColumn: 0,
            endColumn: 1,
          },
        }),
      )
      expect(result.cells).toEqual([
        { row: 0, column: 0, value: 'Old' },
        { row: 0, column: 1, value: 10 },
      ])
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })

  it('opens a blankXlsxBuffer workbook and saves a first edit into it', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-sidecar-blank-'))
    const sourcePath = join(directory, 'blank.xlsx')
    const targetPath = join(directory, 'saved.xlsx')
    await writeFile(sourcePath, await blankXlsxBuffer())
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    let sessionId: string | null = null
    try {
      const opened = openResultSchema.parse(await client.open(sourcePath))
      sessionId = opened.sessionId
      expect(opened.sheets.map((sheet) => sheet.name)).toEqual(['Sheet1'])

      await saveWorkbookViaSidecar({
        client,
        sourcePath,
        targetPath,
        edits: [
          {
            sheetName: 'Sheet1',
            row: 0,
            column: 0,
            writeValue: true,
            cell: { value: 'hello' },
          },
        ],
      })
      const savedZip = await JSZip.loadAsync(await readFile(targetPath))
      const sheet = await savedZip.file('xl/worksheets/sheet1.xml')?.async('text')
      expect(sheet).toContain('hello')
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })

  it('reports pivot-table output ranges in the sheet metadata', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      `<?xml version="1.0"?>
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets><sheet name="Pivot" sheetId="1" r:id="rId1"/></sheets>
      </workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
          Target="worksheets/sheet1.xml"/>
      </Relationships>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<?xml version="1.0"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <dimension ref="A1:F30"/>
        <sheetData><row r="3"><c r="A3" t="inlineStr"><is><t>Row Labels</t></is></c></row></sheetData>
      </worksheet>`,
    )
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId2"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
          Target="../pivotTables/pivotTable1.xml"/>
      </Relationships>`,
    )
    zip.file(
      'xl/pivotTables/pivotTable1.xml',
      `<?xml version="1.0"?>
      <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P1">
        <location ref="A3:C20" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
      </pivotTableDefinition>`,
    )

    const directory = await mkdtemp(join(tmpdir(), 'xlsx-sidecar-pivot-'))
    const path = join(directory, 'pivot.xlsx')
    await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    let sessionId: string | null = null
    try {
      const opened = (await client.open(path)) as {
        sessionId: string
        sheets: { pivotRanges: unknown }[]
      }
      sessionId = opened.sessionId
      expect(opened.sheets[0]?.pivotRanges).toEqual([
        { startRow: 2, startColumn: 0, endRow: 19, endColumn: 2 },
      ])
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })

  it('extracts styles, chart caches, anchors, and embedded image bytes', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-visual-test-'))
    const path = join(directory, 'visuals.xlsx')
    await writeFile(path, await buildVisualFixture())
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    let sessionId: string | null = null
    try {
      const opened = z
        .object({
          sessionId: z.string().uuid(),
          sheets: z.array(
            z
              .object({
                columnWidths: z.array(
                  z.object({
                    startColumn: z.number(),
                    endColumn: z.number(),
                    width: z.number(),
                  }),
                ),
                defaultRowHeight: z.number().nullable(),
              })
              .passthrough(),
          ),
          styles: z.array(
            z
              .object({
                bold: z.boolean(),
                fillColor: z.string().optional(),
              })
              .passthrough(),
          ),
          visuals: z.array(
            z
              .object({
                id: z.string(),
                kind: z.string(),
                anchor: z
                  .object({
                    fromRow: z.number(),
                    fromColumn: z.number(),
                    toRow: z.number(),
                    toColumn: z.number(),
                  })
                  .passthrough(),
                chart: z
                  .object({
                    title: z.string(),
                    chartTypes: z.array(z.string()),
                    series: z.array(
                      z
                        .object({
                          categories: z.array(z.string()),
                          values: z.array(z.number()),
                        })
                        .passthrough(),
                    ),
                  })
                  .optional(),
              })
              .passthrough(),
          ),
        })
        .parse(await client.open(path))
      sessionId = opened.sessionId
      expect(opened.sheets[0]).toMatchObject({
        columnWidths: [{ startColumn: 1, endColumn: 2, width: 20 }],
        defaultRowHeight: 18,
      })
      expect(opened.styles[1]).toMatchObject({ bold: true, fillColor: '#4472C4' })
      expect(opened.visuals).toHaveLength(2)
      expect(opened.visuals[0]).toMatchObject({
        kind: 'chart',
        anchor: { fromRow: 2, fromColumn: 1, toRow: 12, toColumn: 7 },
        chart: {
          title: 'Quarterly revenue',
          chartTypes: ['barChart'],
          series: [{ categories: ['Q1', 'Q2'], values: [12, 18] }],
        },
      })
      const image = opened.visuals[1]
      if (!image) throw new Error('Visual fixture image was not parsed.')
      expect(image).toMatchObject({ kind: 'image' })
      const media = z
        .object({
          mediaType: z.string(),
          base64: z.string(),
        })
        .parse(
          await client.readMedia({
            sessionId,
            visualId: image.id,
          }),
        )
      expect(media.mediaType).toBe('image/png')
      expect(Buffer.from(media.base64, 'base64').subarray(1, 4).toString()).toBe('PNG')

      const range = workbookRangeResultSchema.parse(
        await client.readRange({
          sessionId,
          sheetId: 'sheet-1',
          range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        }),
      )
      expect(range.cells[0]).toMatchObject({ value: 'Styled', styleIndex: 1 })
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })
  it('extracts theme colors, borders, freeze panes, hidden rows, and merges', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-structure-test-'))
    const path = join(directory, 'structure.xlsx')
    await writeFile(path, await buildStructureFixture())
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    let sessionId: string | null = null
    try {
      const opened = z
        .object({
          sessionId: z.string().uuid(),
          sheets: z.array(
            z
              .object({
                rowCount: z.number(),
                columnCount: z.number(),
                columnWidths: z.array(
                  z.object({
                    startColumn: z.number(),
                    endColumn: z.number(),
                    width: z.number().optional(),
                    hidden: z.boolean(),
                    outlineLevel: z.number().optional(),
                    collapsed: z.boolean().optional(),
                    styleIndex: z.number().optional(),
                  }),
                ),
                freeze: z
                  .object({
                    frozenColumns: z.number(),
                    frozenRows: z.number(),
                  })
                  .nullable(),
                hidden: z.boolean(),
                tabColor: z.string().nullable(),
                showGridLines: z.boolean(),
                defaultColumnWidth: z.number().nullable(),
              })
              .passthrough(),
          ),
          styles: z.array(
            z
              .object({
                bold: z.boolean(),
                underline: z.boolean(),
                strikethrough: z.boolean(),
                wrapText: z.boolean(),
              })
              .passthrough(),
          ),
          dxfStyles: z.array(
            z
              .object({
                bold: z.boolean(),
              })
              .passthrough(),
          ),
          definedNames: z.array(
            z
              .object({
                name: z.string(),
                formula: z.string(),
              })
              .strict(),
          ),
        })
        .parse(await client.open(path))
      sessionId = opened.sessionId

      expect(opened.sheets[0]?.freeze).toEqual({ frozenColumns: 1, frozenRows: 2 })
      expect(opened.sheets[0]?.columnWidths).toEqual([
        { startColumn: 1, endColumn: 1, width: 0, hidden: true },
        { startColumn: 2, endColumn: 2, width: 24, hidden: false },
        { startColumn: 3, endColumn: 3, hidden: false, outlineLevel: 2, collapsed: true },
        { startColumn: 4, endColumn: 5, hidden: false, styleIndex: 2 },
      ])
      expect(opened.sheets[0]).toMatchObject({
        hidden: false,
        tabColor: '#92D050',
        showGridLines: false,
        defaultColumnWidth: 12,
      })
      expect(opened.sheets[1]).toMatchObject({ hidden: true })
      expect(opened.definedNames).toEqual([{ name: 'MyRange', formula: 'Structure!$A$1:$B$2' }])
      expect(opened.dxfStyles[0]).toMatchObject({
        bold: true,
        fontColor: '#9C0006',
        fillColor: '#FFC7CE',
        numberFormat: '#,##0.0',
      })
      expect(opened.sheets[0]?.tables).toEqual([
        {
          range: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 3 },
          headerRowCount: 1,
          showRowStripes: true,
          showColumnStripes: false,
          name: 'Table1',
          columns: ['Item', 'B', 'C', 'D'],
          styleName: 'TableStyleMedium2',
          headerFill: '#4472C4',
          headerFontColor: '#FFFFFF',
          stripeFill: '#DAE3F3',
        },
      ])
      expect(opened.styles[1]).toMatchObject({
        borderDiagonal: { style: 'thin', color: '#00B050' },
        diagonalUp: true,
        diagonalDown: false,
      })
      expect(opened.sheets[0]?.comments).toEqual([
        { row: 2, column: 0, author: 'Tester', text: 'note here' },
      ])
      expect(opened.styles[1]).toMatchObject({
        underline: true,
        strikethrough: true,
        wrapText: true,
        fontColor: '#4472C4',
        numberFormat: '0%',
        textRotation: 90,
        borderLeft: { style: 'thin', color: '#FF0000' },
        borderTop: { style: 'medium', color: '#4472C4' },
        borderBottom: { style: 'double' },
      })
      expect(opened.styles[2]).toMatchObject({
        fontColor: '#993366',
        fillColor: '#F8CBAD',
        textRotation: 255,
      })

      let result = workbookRangeResultSchema.parse(
        await client.readRange({
          sessionId,
          sheetId: 'sheet-1',
          range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 },
        }),
      )
      for (let attempt = 0; attempt < 40 && !result.indexingComplete; attempt += 1) {
        await new Promise((resolve) => setTimeout(resolve, 100))
        result = workbookRangeResultSchema.parse(
          await client.readRange({
            sessionId,
            sheetId: 'sheet-1',
            range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 },
          }),
        )
      }
      expect(result.indexingComplete).toBe(true)
      expect(result.merges).toEqual([{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }])
      expect(result.rows).toEqual([
        { row: 0, height: 30, customHeight: true, hidden: false, styleIndex: 2 },
        { row: 1, hidden: true },
        { row: 2, hidden: false, outlineLevel: 1 },
      ])
      expect(result.hyperlinks).toEqual([{ row: 3, column: 0, target: 'https://www.genspark.ai' }])
      const richCell = result.cells.find((cell) => cell.row === 2 && cell.column === 1)
      expect(richCell?.value).toBe('Red plain')
      expect(richCell?.rich).toEqual([
        {
          text: 'Red',
          bold: true,
          italic: false,
          underline: false,
          strikethrough: false,
          color: '#FF0000',
          size: 14,
          family: 'Arial',
        },
        {
          text: ' plain',
          bold: false,
          italic: false,
          underline: false,
          strikethrough: false,
        },
      ])
      expect(result.conditionalRules).toHaveLength(3)
      expect(result.conditionalRules[0]).toMatchObject({
        ruleType: 'cellIs',
        operator: 'greaterThan',
        formulas: ['6'],
        dxfIndex: 0,
        priority: 1,
        ranges: [{ startRow: 1, startColumn: 0, endRow: 2, endColumn: 0 }],
      })
      expect(result.conditionalRules[1]).toMatchObject({
        ruleType: 'dataBar',
        priority: 2,
        cfvos: [{ kind: 'min' }, { kind: 'max' }],
        colors: ['#638EC6'],
        showValue: false,
        negativeColor: '#FF0000',
        negativeSameAsPositive: false,
        gradient: false,
      })
      expect(result.conditionalRules[2]).toMatchObject({
        ruleType: 'iconSet',
        iconSetName: '3Arrows',
        showValue: false,
        cfvos: [
          { kind: 'percent', value: '0' },
          { kind: 'percent', value: '33' },
          { kind: 'percent', value: '67', gte: false },
        ],
      })
      expect(result.autoFilter).toEqual({
        startRow: 0,
        startColumn: 0,
        endRow: 3,
        endColumn: 3,
      })
      expect(result.dataValidations).toEqual([
        {
          ranges: [{ startRow: 1, startColumn: 2, endRow: 2, endColumn: 2 }],
          ruleType: 'list',
          formulas: ['"Yes,No"'],
          allowBlank: true,
          suppressDropdown: false,
          showInputMessage: false,
          showErrorMessage: false,
        },
      ])
      const errorCell = result.cells.find((cell) => cell.row === 3 && cell.column === 1)
      expect(errorCell?.value).toBe('#VALUE!')
      const entityCell = result.cells.find((cell) => cell.row === 1 && cell.column === 3)
      expect(entityCell?.value).toBe('mixed 样式 & more')
      const strayTextCell = result.cells.find((cell) => cell.row === 3 && cell.column === 3)
      expect(strayTextCell?.value).toBe('not-a-number')
      const inlineRichCell = result.cells.find((cell) => cell.row === 3 && cell.column === 2)
      expect(inlineRichCell?.value).toBe('Inline')
      expect(inlineRichCell?.rich).toEqual([
        { text: 'In', bold: true, italic: false, underline: false, strikethrough: false },
        { text: 'line', bold: false, italic: false, underline: false, strikethrough: false },
      ])
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })
})

async function buildStructureFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets>
        <sheet name="Structure" sheetId="1" r:id="rId1"/>
        <sheet name="Secret" sheetId="2" state="hidden" r:id="rId2"/>
      </sheets>
      <definedNames>
        <definedName name="MyRange">Structure!$A$1:$B$2</definedName>
        <definedName name="_xlnm.Print_Area">Structure!$A$1:$D$4</definedName>
      </definedNames>
    </workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet2.xml"/>
    </Relationships>`,
  )
  zip.file(
    'xl/sharedStrings.xml',
    `<?xml version="1.0"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
      <si>
        <r><rPr><b/><color rgb="FFFF0000"/><sz val="14"/><rFont val="Arial"/></rPr><t>Red</t></r>
        <r><t xml:space="preserve"> plain</t></r>
      </si>
    </sst>`,
  )
  zip.file(
    'xl/theme/theme1.xml',
    `<?xml version="1.0"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
      <a:themeElements><a:clrScheme name="Office">
        <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
        <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="44546A"/></a:dk2>
        <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
        <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
        <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
        <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
        <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
        <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
      </a:clrScheme></a:themeElements>
    </a:theme>`,
  )
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="3">
        <font><name val="Calibri"/><sz val="11"/></font>
        <font><u/><strike/><color theme="4"/><name val="Arial"/><sz val="12"/></font>
        <font><color indexed="25"/></font>
      </fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor theme="5" tint="0.6"/></patternFill></fill>
      </fills>
      <borders count="2">
        <border/>
        <border diagonalUp="1">
          <left style="thin"><color rgb="FFFF0000"/></left>
          <right style="thin"/>
          <top style="medium"><color theme="4"/></top>
          <bottom style="double"/>
          <diagonal style="thin"><color rgb="FF00B050"/></diagonal>
        </border>
      </borders>
      <cellXfs count="3">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="9" fontId="1" fillId="0" borderId="1"><alignment horizontal="center" wrapText="1" textRotation="90"/></xf>
        <xf numFmtId="0" fontId="2" fillId="2" borderId="0"><alignment textRotation="255"/></xf>
      </cellXfs>
      <dxfs count="1">
        <dxf><font><b/><color rgb="FF9C0006"/></font><numFmt numFmtId="165" formatCode="#,##0.0"/><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
      </dxfs>
    </styleSheet>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr><tabColor rgb="FF92D050"/></sheetPr>
      <dimension ref="A1:D4"/>
      <sheetViews><sheetView workbookViewId="0" showGridLines="0">
        <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
      </sheetView></sheetViews>
      <sheetFormatPr defaultRowHeight="15" defaultColWidth="12"/>
      <cols>
        <col min="2" max="2" width="0" hidden="1" customWidth="1"/>
        <col min="3" max="3" width="24" customWidth="1"/>
        <col min="4" max="4" outlineLevel="2" collapsed="1"/>
        <col min="5" max="6" style="2"/>
      </cols>
      <sheetData>
        <row r="1" ht="30" customHeight="1" s="2" customFormat="1"><c r="A1" t="inlineStr" s="1"><is><t>Merged</t></is></c></row>
        <row r="2" hidden="1"><c r="A2"><v>5</v></c></row>
        <row r="3" outlineLevel="1"><c r="A3" s="2"><v>7</v></c><c r="B3" t="s"><v>0</v></c></row>
        <row r="4"><c r="A4" t="inlineStr"><is><t>Link</t></is></c><c r="B4" t="e"><v>#VALUE!</v></c><c r="C4" t="inlineStr"><is><r><rPr><b/></rPr><t>In</t></r><r><t>line</t></r></is></c><c r="D4"><v>not-a-number</v></c></row>
        <row r="2"><c r="D2" t="inlineStr"><is><t>mixed &#26679;&#24335; &amp; more</t></is></c></row>
      </sheetData>
      <autoFilter ref="A1:D4"/>
      <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
      <conditionalFormatting sqref="A2:A3">
        <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>6</formula></cfRule>
        <cfRule type="dataBar" priority="2">
          <dataBar showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>
          <extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{TEST-DB-1}</x14:id></ext></extLst>
        </cfRule>
        <cfRule type="iconSet" priority="3">
          <iconSet iconSet="3Arrows" showValue="0"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67" gte="0"/></iconSet>
        </cfRule>
      </conditionalFormatting>
      <dataValidations count="1">
        <dataValidation type="list" allowBlank="1" sqref="C2:C3"><formula1>"Yes,No"</formula1></dataValidation>
      </dataValidations>
      <hyperlinks><hyperlink ref="A4" r:id="rId1"/></hyperlinks>
      <extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"
        xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
        <x14:conditionalFormattings>
          <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:cfRule type="dataBar" id="{TEST-DB-1}">
              <x14:dataBar minLength="0" maxLength="100" gradient="0" negativeBarColorSameAsPositive="0">
                <x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>
                <x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/>
              </x14:dataBar>
            </x14:cfRule>
            <xm:sqref>A2:A3</xm:sqref>
          </x14:conditionalFormatting>
        </x14:conditionalFormattings>
      </ext></extLst>
    </worksheet>`,
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        Target="https://www.genspark.ai" TargetMode="External"/>
      <Relationship Id="rId2"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
        Target="../tables/table1.xml"/>
      <Relationship Id="rId3"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        Target="../comments1.xml"/>
    </Relationships>`,
  )
  zip.file(
    'xl/tables/table1.xml',
    `<?xml version="1.0"?>
    <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      id="1" name="Table1" displayName="Table1" ref="A1:D4" headerRowCount="1">
      <tableColumns count="4">
        <tableColumn id="1" name="Item"/><tableColumn id="2" name="B"/>
        <tableColumn id="3" name="C"/><tableColumn id="4" name="D"/>
      </tableColumns>
      <tableStyleInfo name="TableStyleMedium2" showRowStripes="1" showColumnStripes="0"/>
    </table>`,
  )
  zip.file(
    'xl/comments1.xml',
    `<?xml version="1.0"?>
    <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <authors><author>Tester</author></authors>
      <commentList>
        <comment ref="A3" authorId="0"><text><r><t>note here</t></r></text></comment>
      </commentList>
    </comments>`,
  )
  zip.file(
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:A1"/>
      <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
    </worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}

async function buildVisualFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="Visuals" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.xml"/>
    </Relationships>`,
  )
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="2"><font><name val="Calibri"/><sz val="11"/></font><font><b/><name val="Arial"/><sz val="12"/></font></fonts>
      <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0"/></cellXfs>
    </styleSheet>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <dimension ref="A1:H20"/><sheetFormatPr defaultRowHeight="18"/>
      <cols><col min="2" max="3" width="20" customWidth="1"/></cols>
      <sheetData><row r="1"><c r="A1" t="inlineStr" s="1"><is><t>Styled</t></is></c></row></sheetData>
      <drawing r:id="rId1"/>
    </worksheet>`,
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
        Target="../drawings/drawing1.xml"/>
    </Relationships>`,
  )
  zip.file(
    'xl/drawings/drawing1.xml',
    `<?xml version="1.0"?>
    <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Revenue chart"/></xdr:nvGraphicFramePr>
          <a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic>
        </xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>
      <xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>13</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId2"/></xdr:blipFill></xdr:pic>
        <xdr:clientData/></xdr:twoCellAnchor>
    </xdr:wsDr>`,
  )
  zip.file(
    'xl/drawings/_rels/drawing1.xml.rels',
    `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
    </Relationships>`,
  )
  zip.file(
    'xl/charts/chart1.xml',
    `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Quarterly revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:formatCode>0</c:formatCode><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>18</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:barChart></c:plotArea></c:chart>
    </c:chartSpace>`,
  )
  zip.file(
    'xl/media/image1.png',
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
      'base64',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}
