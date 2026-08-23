import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  assembleWithJsZip,
  assertOnlyTouchedEntriesChanged,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '../src/gateway/xlsx-gateway'
import type { CellEdit, SheetVisualAddition } from '../src/gateway/xlsx-gateway'
import type { SheetEditPlan } from '../src/gateway/xlsx-sheets'
import { buildEditFixture, buildSheetsFixture } from './fixture-builder'

const ANCHOR = {
  fromRow: 1,
  fromColumn: 5,
  fromRowOffset: 0,
  fromColumnOffset: 0,
  toRow: 16,
  toColumn: 12,
  toRowOffset: 0,
  toColumnOffset: 0,
}

function chartAddition(overrides: Partial<SheetVisualAddition['chart']> = {}): SheetVisualAddition {
  return {
    sheetName: 'Data',
    anchor: ANCHOR,
    chart: {
      chartType: 'column',
      title: 'Users by <Country>',
      series: [
        {
          name: 'Users',
          categories: ['US', 'KR', 'BR'],
          values: [8, 6, 3],
          valuesRef: "'Data'!$B$2:$B$4",
          categoriesRef: "'Data'!$A$2:$A$4",
        },
      ],
      ...overrides,
    },
  }
}

async function planWith(additions: SheetVisualAddition[], buffer?: Buffer) {
  const source = await createBufferEntrySource(buffer ?? (await buildEditFixture()))
  return planCellEditsToXlsx(source, [], [], [], undefined, [], [], [], [], [], null, additions)
}

describe('visual additions: fresh sheet without a drawing', () => {
  it('creates chart + drawing parts and registers every relationship', async () => {
    const plan = await planWith([chartAddition()])

    // fixture already ships xl/charts/chart1.xml → new chart takes index 2
    const chartXml = plan.added.get('xl/charts/chart2.xml')
    expect(chartXml).toBeDefined()
    expect(chartXml).toContain('<a:t>Users by &lt;Country&gt;</a:t>')
    expect(chartXml).toContain('<c:barDir val="col"/>')
    expect(chartXml).toContain("<c:f>'Data'!$B$2:$B$4</c:f>")
    expect(chartXml).toContain("<c:f>'Data'!$A$2:$A$4</c:f>")
    expect(chartXml).toContain('<c:pt idx="2"><c:v>3</c:v></c:pt>')
    expect(chartXml).toContain('<c:pt idx="0"><c:v>US</c:v></c:pt>')

    const drawingXml = plan.added.get('xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<xdr:twoCellAnchor>')
    expect(drawingXml).toContain('<xdr:col>5</xdr:col>')
    expect(drawingXml).toContain('r:id="rId1"')

    const drawingRels = plan.added.get('xl/drawings/_rels/drawing1.xml.rels')
    expect(drawingRels).toContain('relationships/chart')
    expect(drawingRels).toContain('Target="../charts/chart2.xml"')

    const sheetRels = plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(sheetRels).toContain('relationships/drawing')
    expect(sheetRels).toContain('Target="../drawings/drawing1.xml"')

    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<drawing r:id="rId1"/>')
    expect(worksheet).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    )

    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain('PartName="/xl/charts/chart2.xml"')
    expect(contentTypes).toContain('drawingml.chart+xml')
    expect(contentTypes).toContain('PartName="/xl/drawings/drawing1.xml"')
  })

  it('pie charts omit axes and enable the legend', async () => {
    const plan = await planWith([chartAddition({ chartType: 'pie' })])
    const chartXml = plan.added.get('xl/charts/chart2.xml')
    expect(chartXml).toContain('<c:pieChart>')
    expect(chartXml).not.toContain('<c:catAx>')
    expect(chartXml).toContain('<c:legend>')
  })

  it('falls back to literal caches when range refs are absent', async () => {
    const plan = await planWith([
      chartAddition({
        series: [{ name: 'S', categories: ['a'], values: [1] }],
      }),
    ])
    const chartXml = plan.added.get('xl/charts/chart2.xml')
    expect(chartXml).toContain('<c:numLit>')
    expect(chartXml).toContain('<c:strLit>')
    expect(chartXml).not.toContain('<c:numRef>')
  })

  it('two additions share one drawing and take sequential chart parts', async () => {
    const plan = await planWith([chartAddition(), chartAddition({ chartType: 'line' })])
    expect(plan.added.get('xl/charts/chart2.xml')).toBeDefined()
    expect(plan.added.get('xl/charts/chart3.xml')).toContain('<c:lineChart>')
    const drawingXml = plan.added.get('xl/drawings/drawing1.xml')
    expect(drawingXml?.match(/<xdr:twoCellAnchor>/g)).toHaveLength(2)
    expect(plan.added.get('xl/drawings/drawing2.xml')).toBeUndefined()
    const drawingRels = plan.added.get('xl/drawings/_rels/drawing1.xml.rels')
    expect(drawingRels).toContain('Id="rId1"')
    expect(drawingRels).toContain('Id="rId2"')
  })
})

describe('visual additions: shapes and text boxes', () => {
  it('writes a shape anchor with fill, no chart part, no drawing rels', async () => {
    const plan = await planWith([
      {
        sheetName: 'Data',
        anchor: ANCHOR,
        shape: { shapeType: 'hexagon', fillColor: '#ddebf7' },
      },
    ])
    const drawingXml = plan.added.get('xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<a:prstGeom prst="hexagon">')
    expect(drawingXml).toContain('<a:srgbClr val="DDEBF7"/>')
    expect(drawingXml).not.toContain('txBox')
    expect(drawingXml).not.toContain('<xdr:txBody>')
    expect(plan.added.get('xl/charts/chart2.xml')).toBeUndefined()
    expect(plan.added.get('xl/drawings/_rels/drawing1.xml.rels')).toBeUndefined()
    // drawing part still registers in content types and the sheet rels
    expect(plan.replaced.get('[Content_Types].xml')).toContain(
      'PartName="/xl/drawings/drawing1.xml"',
    )
    expect(plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')).toContain('relationships/drawing')
  })

  it('text boxes carry txBox, an outline, and the text body', async () => {
    const plan = await planWith([
      {
        sheetName: 'Data',
        anchor: ANCHOR,
        shape: { shapeType: 'rect', fillColor: '#FFFFFF', text: 'Hello <world>', isTextBox: true },
      },
    ])
    const drawingXml = plan.added.get('xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<xdr:cNvSpPr txBox="1"/>')
    expect(drawingXml).toContain('name="TextBox 2"')
    expect(drawingXml).toContain('<a:t>Hello &lt;world&gt;</a:t>')
    expect(drawingXml).toContain('<a:ln w="9525">')
  })

  it('a chart and a shape share the sheet drawing', async () => {
    const plan = await planWith([
      chartAddition(),
      { sheetName: 'Data', anchor: ANCHOR, shape: { shapeType: 'rect' } },
    ])
    const drawingXml = plan.added.get('xl/drawings/drawing1.xml')
    expect(drawingXml?.match(/<xdr:twoCellAnchor>/g)).toHaveLength(2)
    expect(drawingXml).toContain('<xdr:graphicFrame')
    expect(drawingXml).toContain('<xdr:sp macro=""')
    expect(drawingXml).toContain('<a:noFill/>')
  })
})

describe('visual additions: pictures', () => {
  const PNG_BASE64 = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]).toString(
    'base64',
  )

  it('writes media bytes, a Default content type, and a pic anchor', async () => {
    const plan = await planWith([
      {
        sheetName: 'Data',
        anchor: ANCHOR,
        image: { mediaType: 'image/png', base64: PNG_BASE64 },
      },
    ])
    const media = plan.addedBinary.get('xl/media/image1.png')
    expect(media).toBeDefined()
    expect(Buffer.from(media!).subarray(0, 4)).toEqual(Buffer.from([0x89, 0x50, 0x4e, 0x47]))
    expect(plan.addedEntries).toContain('xl/media/image1.png')

    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain('Extension="png"')
    expect(contentTypes).toContain('ContentType="image/png"')

    const drawingXml = plan.added.get('xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<xdr:pic>')
    expect(drawingXml).toContain('r:embed="rId1"')
    const drawingRels = plan.added.get('xl/drawings/_rels/drawing1.xml.rels')
    expect(drawingRels).toContain('relationships/image')
    expect(drawingRels).toContain('Target="../media/image1.png"')
  })

  it('does not duplicate an existing Default extension declaration', async () => {
    const plan = await planWith([
      { sheetName: 'Data', anchor: ANCHOR, image: { mediaType: 'image/png', base64: PNG_BASE64 } },
      { sheetName: 'Data', anchor: ANCHOR, image: { mediaType: 'image/png', base64: PNG_BASE64 } },
    ])
    expect(plan.addedBinary.get('xl/media/image2.png')).toBeDefined()
    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes?.match(/Extension="png"/g)).toHaveLength(1)
  })
})

describe('visual additions: sheet with an existing drawing', () => {
  const EMPTY_SELF_CLOSED_DRAWING =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'

  async function buildDrawingFixture(drawingXml?: string): Promise<Buffer> {
    const zip = await JSZip.loadAsync(await buildEditFixture())
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    zip.file(
      'xl/worksheets/sheet1.xml',
      worksheet
        .replace(
          '<worksheet ',
          '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
        )
        .replace('</worksheet>', '<drawing r:id="rId9"/></worksheet>'),
    )
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/drawings/drawing1.xml',
      drawingXml ??
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
          ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
          '<xdr:twoCellAnchor><xdr:graphicFrame><xdr:nvGraphicFramePr>' +
          '<xdr:cNvPr id="3" name="Chart 3"/></xdr:nvGraphicFramePr></xdr:graphicFrame></xdr:twoCellAnchor>' +
          '</xdr:wsDr>',
    )
    zip.file(
      'xl/drawings/_rels/drawing1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>' +
        '</Relationships>',
    )
    return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  }

  it('appends the anchor to the existing drawing with fresh ids', async () => {
    const plan = await planWith([chartAddition()], await buildDrawingFixture())

    expect(plan.added.get('xl/drawings/drawing1.xml')).toBeUndefined()
    const drawingXml = plan.replaced.get('xl/drawings/drawing1.xml')
    expect(drawingXml?.match(/<xdr:twoCellAnchor>/g)).toHaveLength(2)
    // next shape id after the existing id=3
    expect(drawingXml).toContain('<xdr:cNvPr id="4" name="Chart 4"/>')
    // chart rel appended after the existing rId1
    expect(drawingXml).toContain('r:id="rId2"')
    const drawingRels = plan.replaced.get('xl/drawings/_rels/drawing1.xml.rels')
    expect(drawingRels).toContain('Id="rId2"')
    expect(drawingRels).toContain('Target="../charts/chart2.xml"')

    // worksheet keeps its single existing <drawing> element
    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheet ?? '').not.toContain('rId10')
  })

  it('expands a self-closed empty wsDr root (Google Sheets exports)', async () => {
    const plan = await planWith(
      [chartAddition()],
      await buildDrawingFixture(EMPTY_SELF_CLOSED_DRAWING),
    )
    const drawingXml = plan.replaced.get('xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<xdr:twoCellAnchor>')
    expect(drawingXml).toMatch(/<\/xdr:wsDr>$/)
    expect(drawingXml).not.toContain('/><xdr:twoCellAnchor>')
    expect(drawingXml).toContain('<xdr:graphicFrame')
  })

  it('declares missing namespaces on the inserted anchor', async () => {
    const drawing =
      '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>'
    const plan = await planWith([chartAddition()], await buildDrawingFixture(drawing))
    const drawingXml = plan.replaced.get('xl/drawings/drawing1.xml')
    expect(drawingXml).toContain(
      '<xdr:twoCellAnchor' +
        ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    )
    expect(drawingXml).toMatch(/<\/wsDr>$/)
  })
})

describe('visual additions: sheets created in the same save', () => {
  const PNG_BASE64 = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]).toString(
    'base64',
  )

  function sheetPlanWith(additions: SheetEditPlan['additions']): SheetEditPlan {
    return {
      renames: [],
      additions,
      removals: [],
      order: ['Data', 'Other', 'My Sheet', ...additions.map((addition) => addition.name)],
    }
  }

  async function saveWithSheetPlan(
    additions: SheetVisualAddition[],
    sheetPlan: SheetEditPlan,
    edits: CellEdit[] = [],
  ) {
    const source = await buildSheetsFixture()
    const plan = await planCellEditsToXlsx(
      await createBufferEntrySource(source),
      edits,
      [],
      [],
      sheetPlan,
      [],
      [],
      [],
      [],
      [],
      null,
      additions,
    )
    const mutation = await assembleWithJsZip(source, plan)
    assertOnlyTouchedEntriesChanged(mutation)
    return mutation
  }

  async function entryText(buffer: Buffer, path: string): Promise<string> {
    const zip = await JSZip.loadAsync(buffer)
    return zip.file(path)?.async('text') ?? Promise.resolve('')
  }

  it('chart on an added sheet: parts, rels, and overrides land on the new part', async () => {
    const mutation = await saveWithSheetPlan(
      [{ ...chartAddition(), sheetName: 'Extra' }],
      sheetPlanWith([{ name: 'Extra' }]),
      [{ sheetName: 'Extra', row: 0, column: 0, writeValue: true, cell: { value: 3 } }],
    )

    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).toContain('<sheet name="Extra" sheetId="4" r:id="rId6"/>')
    const workbookRels = await entryText(mutation.buffer, 'xl/_rels/workbook.xml.rels')
    expect(workbookRels).toContain('Id="rId6"')
    expect(workbookRels).toContain('Target="worksheets/sheet4.xml"')

    // the edit and the drawing element both survive on the minimal new part,
    // with <drawing> after sheetData per CT_Worksheet order
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet4.xml')
    expect(worksheet).toContain('<c r="A1"><v>3</v></c>')
    expect(worksheet).toMatch(/<\/sheetData><drawing r:id="rId1"\/><\/worksheet>$/)

    const sheetRels = await entryText(mutation.buffer, 'xl/worksheets/_rels/sheet4.xml.rels')
    expect(sheetRels).toContain('relationships/drawing')
    expect(sheetRels).toContain('Target="../drawings/drawing1.xml"')

    const drawingXml = await entryText(mutation.buffer, 'xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<xdr:twoCellAnchor>')
    expect(drawingXml).toContain('<xdr:graphicFrame')
    const drawingRels = await entryText(mutation.buffer, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(drawingRels).toContain('relationships/chart')
    expect(drawingRels).toContain('Target="../charts/chart2.xml"')
    const chartXml = await entryText(mutation.buffer, 'xl/charts/chart2.xml')
    expect(chartXml).toContain('<c:barDir val="col"/>')

    const contentTypes = await entryText(mutation.buffer, '[Content_Types].xml')
    expect(contentTypes).toContain('PartName="/xl/worksheets/sheet4.xml"')
    expect(contentTypes).toContain('PartName="/xl/drawings/drawing1.xml"')
    expect(contentTypes).toContain('PartName="/xl/charts/chart2.xml"')
  })

  it('shape on an added sheet without cell edits', async () => {
    const mutation = await saveWithSheetPlan(
      [
        {
          sheetName: 'Extra',
          anchor: ANCHOR,
          shape: { shapeType: 'hexagon', fillColor: '#ddebf7' },
        },
      ],
      sheetPlanWith([{ name: 'Extra' }]),
    )
    const worksheet = await entryText(mutation.buffer, 'xl/worksheets/sheet4.xml')
    expect(worksheet).toMatch(/<sheetData\/><drawing r:id="rId1"\/><\/worksheet>$/)
    const drawingXml = await entryText(mutation.buffer, 'xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<a:prstGeom prst="hexagon">')
    expect(drawingXml).toContain('<a:srgbClr val="DDEBF7"/>')
    expect(await entryText(mutation.buffer, '[Content_Types].xml')).toContain(
      'PartName="/xl/drawings/drawing1.xml"',
    )
  })

  it('image on an added sheet writes media bytes and the Default extension', async () => {
    const mutation = await saveWithSheetPlan(
      [
        {
          sheetName: 'Extra',
          anchor: ANCHOR,
          image: { mediaType: 'image/png', base64: PNG_BASE64 },
        },
      ],
      sheetPlanWith([{ name: 'Extra' }]),
    )
    expect(mutation.addedEntries).toContain('xl/media/image1.png')
    const zip = await JSZip.loadAsync(mutation.buffer)
    const media = await zip.file('xl/media/image1.png')!.async('nodebuffer')
    expect(media.subarray(0, 4)).toEqual(Buffer.from([0x89, 0x50, 0x4e, 0x47]))
    const contentTypes = await entryText(mutation.buffer, '[Content_Types].xml')
    expect(contentTypes).toContain('Extension="png"')
    const drawingXml = await entryText(mutation.buffer, 'xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<xdr:pic>')
    expect(drawingXml).toContain('r:embed="rId1"')
    const drawingRels = await entryText(mutation.buffer, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(drawingRels).toContain('Target="../media/image1.png"')
    expect(await entryText(mutation.buffer, 'xl/worksheets/sheet4.xml')).toContain(
      '<drawing r:id="rId1"/>',
    )
  })

  it('chart on a duplicated sheet: clone keeps its content and gains the drawing', async () => {
    const mutation = await saveWithSheetPlan([{ ...chartAddition(), sheetName: 'Data Copy' }], {
      renames: [],
      additions: [{ name: 'Data Copy', sourceSheetName: 'Data' }],
      removals: [],
      order: ['Data', 'Data Copy', 'Other', 'My Sheet'],
    })
    const copy = await entryText(mutation.buffer, 'xl/worksheets/sheet4.xml')
    expect(copy).toContain("<f>'My Sheet'!A1*2</f>")
    expect(copy).toContain('<drawing r:id="rId1"/>')
    // the source part stays byte-untouched
    expect(mutation.touchedEntries).not.toContain('xl/worksheets/sheet1.xml')
    expect(await entryText(mutation.buffer, 'xl/worksheets/sheet1.xml')).not.toContain('<drawing')
    const sheetRels = await entryText(mutation.buffer, 'xl/worksheets/_rels/sheet4.xml.rels')
    expect(sheetRels).toContain('Target="../drawings/drawing1.xml"')
    const drawingRels = await entryText(mutation.buffer, 'xl/drawings/_rels/drawing1.xml.rels')
    expect(drawingRels).toContain('Target="../charts/chart2.xml"')
    expect(await entryText(mutation.buffer, 'xl/charts/chart2.xml')).toContain(
      '<c:barDir val="col"/>',
    )
    const workbook = await entryText(mutation.buffer, 'xl/workbook.xml')
    expect(workbook).toContain('<sheet name="Data Copy" sheetId="4" r:id="rId6"/>')
  })
})
