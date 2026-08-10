import JSZip from 'jszip'

export async function buildCompatibilityFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', contentTypes)
  zip.file('_rels/.rels', packageRelationships)
  zip.file('xl/workbook.xml', workbook)
  zip.file('xl/_rels/workbook.xml.rels', workbookRelationships)
  zip.file('xl/worksheets/sheet1.xml', worksheet)
  zip.file('xl/styles.xml', styles)
  zip.file('customXml/item1.xml', '<compatibility-marker value="must-survive"/>')
  zip.file(
    'xl/charts/chart1.xml',
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/// Exercises the save path: shared strings, styled cells, sparse rows, a
/// self-closing row, a cached formula, and parts that must survive untouched.
export async function buildEditFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', editContentTypes)
  zip.file('_rels/.rels', packageRelationships)
  zip.file('xl/workbook.xml', editWorkbook)
  zip.file('xl/_rels/workbook.xml.rels', editWorkbookRelationships)
  zip.file('xl/worksheets/sheet1.xml', editWorksheet)
  zip.file('xl/styles.xml', editStyles)
  zip.file('xl/sharedStrings.xml', editSharedStrings)
  zip.file('customXml/item1.xml', '<compatibility-marker value="must-survive"/>')
  zip.file(
    'xl/charts/chart1.xml',
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const editContentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`

const editWorkbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <definedNames><definedName name="Total">Data!$C$1</definedName></definedNames>
</workbook>`

const editWorkbookRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`

const editWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="C1"><v>5</v></c></row>
    <row r="3"><c r="B3"><f>SUM(C1:C2)</f><v>5</v></c></row>
    <row r="5" ht="20" customHeight="1"/>
  </sheetData>
</worksheet>`

const editSharedStrings = `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Hello</t></si></sst>`

const editStyles = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font/><font><b/></font></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>
</styleSheet>`

/// Exercises structural row/column shifts: formulas (absolute/relative/range/
/// cross-sheet/string literal), merges, cols, CF, DV, hyperlinks, calcChain.
export async function buildStructureFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', structureContentTypes)
  zip.file('_rels/.rels', packageRelationships)
  zip.file('xl/workbook.xml', structureWorkbook)
  zip.file('xl/_rels/workbook.xml.rels', structureWorkbookRelationships)
  zip.file('xl/worksheets/sheet1.xml', structureWorksheet)
  zip.file('xl/worksheets/sheet2.xml', structureOtherWorksheet)
  zip.file('xl/styles.xml', styles)
  zip.file(
    'xl/calcChain.xml',
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="D2" i="1"/></calcChain>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const structureContentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`

const structureWorkbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>
</workbook>`

const structureWorkbookRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`

const structureWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D10"/>
  <cols><col min="2" max="3" width="20" customWidth="1"/></cols>
  <sheetData>
    <row r="1" spans="1:4"><c r="A1"><v>1</v></c><c r="C1"><v>7</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c><c r="D2"><f>SUM(A1:A2)</f><v>3</v></c></row>
    <row r="4"><c r="B4"><f>$A$2+Other!B9+SUM(B:B)&amp;"A1"</f><v>0</v></c></row>
    <row r="10"><c r="A10"><v>10</v></c></row>
  </sheetData>
  <mergeCells count="2"><mergeCell ref="A5:B6"/><mergeCell ref="C1:D1"/></mergeCells>
  <conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="1"><formula>$A2&gt;5</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="list" sqref="B2:B3"><formula1>"a,b"</formula1></dataValidation></dataValidations>
</worksheet>`

const structureOtherWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`

/// Exercises sheet rename/add/remove: three sheets (one with a quoted name),
/// cross-sheet formulas, an internal hyperlink anchor, a chart series
/// reference, scoped and global defined names, an active tab, and calcChain.
export async function buildSheetsFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', sheetsContentTypes)
  zip.file('_rels/.rels', packageRelationships)
  zip.file('xl/workbook.xml', sheetsWorkbook)
  zip.file('xl/_rels/workbook.xml.rels', sheetsWorkbookRelationships)
  zip.file('xl/worksheets/sheet1.xml', sheetsDataWorksheet)
  zip.file('xl/worksheets/sheet2.xml', sheetsOtherWorksheet)
  zip.file('xl/worksheets/sheet3.xml', sheetsMyWorksheet)
  zip.file('xl/styles.xml', styles)
  zip.file('xl/charts/chart1.xml', sheetsChart)
  zip.file(
    'xl/calcChain.xml',
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A2" i="1"/></calcChain>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const sheetsContentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`

const sheetsWorkbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="2"/></bookViews>
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/><sheet name="My Sheet" sheetId="3" r:id="rId3"/></sheets>
  <definedNames><definedName name="GlobalTotal">Data!$A$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="1">Other!$A$1:$B$2</definedName><definedName name="LocalMy" localSheetId="2">'My Sheet'!$A$1</definedName></definedNames>
</workbook>`

const sheetsWorkbookRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`

const sheetsDataWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2"><c r="A2"><f>'My Sheet'!A1*2</f><v>10</v></c></row>
  </sheetData>
</worksheet>`

const sheetsOtherWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f>Data!A1+SUM('My Sheet'!A1:A2)</f><v>6</v></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>go</t></is></c></row>
  </sheetData>
  <hyperlinks><hyperlink ref="A2" location="Data!A1" display="go"/></hyperlinks>
</worksheet>`

const sheetsMyWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>5</v></c></row>
  </sheetData>
</worksheet>`

const sheetsChart = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Data!$A$1:$A$2</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`

/// Exercises sheet removal with owned satellite parts: "Deco" carries a
/// drawing (image + chart with a color style), comments, legacy VML, and a
/// table. Options add a second drawing on the surviving sheet sharing the
/// image, a pivot-table relationship, a surviving structured reference into
/// the table, or a defined name scoped to "Deco" that uses the table.
export async function buildSatelliteSheetFixture(
  options: {
    sharedImage?: boolean
    pivot?: boolean
    tableRef?: boolean
    scopedTableName?: boolean
  } = {},
): Promise<Buffer> {
  const zip = new JSZip()
  const keepFormula = options.tableRef ? '<c r="B1"><f>SUM(DecoTable[Amt])</f><v>3</v></c>' : ''
  const keepDrawing = options.sharedImage ? '<drawing r:id="rId1"/>' : ''
  // A name scoped to Deco (localSheetId 1) dies with the sheet, so its
  // structured reference into DecoTable must not block the removal.
  const definedNames = options.scopedTableName
    ? '<definedNames><definedName name="DecoRef" localSheetId="1">SUM(DecoTable[Amt])</definedName></definedNames>'
    : ''
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  ${options.sharedImage ? '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : ''}
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`,
  )
  zip.file('_rels/.rels', packageRelationships)
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Keep" sheetId="1" r:id="rId1"/><sheet name="Deco" sheetId="2" r:id="rId2"/></sheets>
  ${definedNames}
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><v>1</v></c>${keepFormula}</row></sheetData>
  ${keepDrawing}
</worksheet>`,
  )
  if (options.sharedImage) {
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>
</Relationships>`,
    )
    zip.file(
      'xl/drawings/drawing2.xml',
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
    )
    zip.file(
      'xl/drawings/_rels/drawing2.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`,
    )
  }
  zip.file(
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><v>5</v></c></row><row r="2"><c r="A2"><v>7</v></c></row></sheetData>
  <drawing r:id="rId1"/>
  <legacyDrawing r:id="rId3"/>
  <tableParts count="1"><tablePart r:id="rId4"/></tableParts>
</worksheet>`,
  )
  zip.file(
    'xl/worksheets/_rels/sheet2.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  ${options.pivot ? '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>' : ''}
</Relationships>`,
  )
  zip.file(
    'xl/drawings/drawing1.xml',
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
  )
  zip.file(
    'xl/drawings/_rels/drawing1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/charts/chart1.xml',
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Deco!$A$1:$A$2</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
  )
  zip.file(
    'xl/charts/_rels/chart1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/charts/colors1.xml',
    '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"/>',
  )
  zip.file('xl/comments1.xml', kitchenSinkComments)
  zip.file('xl/drawings/vmlDrawing1.vml', kitchenSinkVml)
  zip.file(
    'xl/tables/table1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="DecoTable" displayName="DecoTable" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="Amt"/></tableColumns></table>`,
  )
  if (options.pivot) {
    zip.file(
      'xl/pivotTables/pivotTable1.xml',
      '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
    )
  }
  zip.file('xl/media/image1.png', kitchenSinkPng)
  zip.file('xl/styles.xml', styles)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/// Exercises preservation of exotic and binary parts: theme, doc properties,
/// comments + VML, a table part, frozen panes, autoFilter, PNG media, and
/// printer settings. Only sheet1.xml may change when a cell is edited.
export async function buildKitchenSinkFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', kitchenSinkContentTypes)
  zip.file('_rels/.rels', kitchenSinkPackageRelationships)
  zip.file('docProps/core.xml', kitchenSinkCore)
  zip.file('docProps/app.xml', kitchenSinkApp)
  zip.file('xl/workbook.xml', kitchenSinkWorkbook)
  zip.file('xl/_rels/workbook.xml.rels', kitchenSinkWorkbookRelationships)
  zip.file('xl/worksheets/sheet1.xml', kitchenSinkWorksheet)
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', kitchenSinkWorksheetRelationships)
  zip.file('xl/styles.xml', styles)
  zip.file('xl/theme/theme1.xml', kitchenSinkTheme)
  zip.file('xl/tables/table1.xml', kitchenSinkTable)
  zip.file('xl/comments1.xml', kitchenSinkComments)
  zip.file('xl/drawings/vmlDrawing1.vml', kitchenSinkVml)
  zip.file('xl/media/image1.png', kitchenSinkPng)
  zip.file(
    'xl/printerSettings/printerSettings1.bin',
    Buffer.from([0x01, 0x02, 0x03, 0xff, 0x00, 0x7f]),
  )
  zip.file('customXml/item1.xml', '<compatibility-marker value="must-survive"/>')
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const kitchenSinkContentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`

const kitchenSinkPackageRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

const kitchenSinkCore = `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>fixture</dc:creator></cp:coreProperties>`

const kitchenSinkApp = `<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>GenOffice Fixture</Application></Properties>`

const kitchenSinkWorkbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

const kitchenSinkWorkbookRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`

const kitchenSinkWorksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/></sheetView></sheetViews>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>hdr</t></is></c></row>
    <row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>3</v></c></row>
  </sheetData>
  <autoFilter ref="A1:B2"/>
  <legacyDrawing r:id="rId3"/>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`

const kitchenSinkWorksheetRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`

const kitchenSinkTheme = `<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture"><a:themeElements/></a:theme>`

const kitchenSinkTable = `<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="hdr"/></tableColumns></table>`

const kitchenSinkComments = `<?xml version="1.0" encoding="UTF-8"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>fixture</author></authors><commentList><comment ref="B1" authorId="0"><text><t>note</t></text></comment></commentList></comments>`

const kitchenSinkVml = `<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="_x0000_s1025" type="#_x0000_t202" style="visibility:hidden"/></xml>`

const kitchenSinkPng = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
])

const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`

const packageRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const workbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

const workbookRelationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

const worksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Old</t></is></c><c r="B1"><v>10</v></c></row>
  </sheetData>
</worksheet>`

const styles = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>
</styleSheet>`
