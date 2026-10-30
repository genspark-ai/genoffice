import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { createBufferEntrySource } from '../src/gateway/xlsx-gateway'
import {
  planExternalLinkBreak,
  transformChartForLinkBreak,
  transformSheetForLinkBreak,
  transformWorkbookForLinkBreak,
} from '../src/gateway/xlsx-external-links'

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
const LINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'

/// A workbook buffer with three external links, one worksheet and one chart.
/// The sheet mixes the shapes Excel writes: a self-closing styled cell, plain
/// numeric and string caches, and a conditional-format rule.
async function buildWorkbook(): Promise<Buffer> {
  const zip = new JSZip()
  const overrides = [1, 2, 3]
    .map(
      (n) =>
        '<Override PartName="/xl/externalLinks/externalLink' +
        n +
        '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>',
    )
    .join('')
  zip.file(
    '[Content_Types].xml',
    DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
      overrides +
      '</Types>',
  )
  zip.file(
    'xl/workbook.xml',
    DECL +
      '<workbook ' +
      NS +
      ' ' +
      R_NS +
      '><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
      '<definedNames><definedName name="Renewal">[3]Sheet1!$A$1</definedName></definedNames>' +
      '<externalReferences><externalReference r:id="rId4"/><externalReference r:id="rId5"/>' +
      '<externalReference r:id="rId6"/></externalReferences></workbook>',
  )
  const externalRels = [1, 2, 3]
    .map(
      (n) =>
        '<Relationship Id="rId' +
        (n + 3) +
        '" Type="' +
        LINK_TYPE +
        '" Target="externalLinks/externalLink' +
        n +
        '.xml"/>',
    )
    .join('')
  zip.file(
    'xl/_rels/workbook.xml.rels',
    DECL +
      '<Relationships ' +
      REL_NS +
      '><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>' +
      externalRels +
      '</Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    DECL +
      '<worksheet ' +
      NS +
      '><sheetData>' +
      '<row r="1"><c r="A1" s="2"/>' +
      '<c r="B1"><f>[2]Sheet1!A1*2</f><v>10</v></c>' +
      '<c r="C1" s="4" t="str"><f>[3]Sheet1!A1&amp;"!"</f><v>x!</v></c></row>' +
      '</sheetData>' +
      '<conditionalFormatting sqref="B1"><cfRule type="expression" priority="1">' +
      '<formula>[2]Sheet1!A1&gt;5</formula></cfRule></conditionalFormatting>' +
      '</worksheet>',
  )
  zip.file(
    'xl/charts/chart1.xml',
    DECL +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef>' +
      '<c:f>[3]Sheet1!$A$1:$A$3</c:f>' +
      '</c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('transformSheetForLinkBreak', () => {
  it('does not swallow a self-closing styled cell printing the next formula', () => {
    const xml =
      '<worksheet><sheetData><row r="1"><c r="A1" s="2"/>' +
      '<c r="B1"><f>[1]Sheet1!A1</f><v>5</v></c></row></sheetData></worksheet>'
    const out = transformSheetForLinkBreak(xml, [1])
    expect(out).toContain('<c r="A1" s="2"/>')
    // the formula became a literal and the empty styled cell is untouched
    expect(out).toContain('<c r="B1"><v>5</v></c>')
    expect(out).not.toContain('<f>')
  })

  it('keeps the literal cell attributes and never emits an empty <v>', () => {
    const xml =
      '<worksheet><sheetData><row r="1">' +
      '<c r="A1" s="3" vm="1"><f>[1]Sheet1!A1</f><v>5</v></c>' +
      '<c r="B1" s="4" t="str"><f>[1]Sheet1!A1</f><v>gone!</v></c>' +
      '<c r="C1" s="5"><f>[1]Sheet1!A1</f></c>' +
      '<c r="D1" s="6"><f>[1]Sheet1!A1</f><v></v></c>' +
      '</row></sheetData></worksheet>'
    const out = transformSheetForLinkBreak(xml, [1])
    // style and value metadata survive, the cache type goes with the formula
    expect(out).toContain('<c r="A1" s="3" vm="1"><v>5</v></c>')
    // a formula-string result is an inline string, not a bare <v>
    expect(out).toContain(
      '<c r="B1" s="4" t="inlineStr"><is><t xml:space="preserve">gone!</t></is></c>',
    )
    // no cached value: the broken link's truth is #REF!, not an empty numeric cache
    expect(out).toContain('<c r="C1" s="5" t="e"><v>#REF!</v></c>')
    expect(out).toContain('<c r="D1" s="6" t="e"><v>#REF!</v></c>')
    expect(out).not.toContain('<c r="A1" s="3"><v></v></c>')
  })

  it('drops the dynamic-array mark from an array formula it turns into a literal', () => {
    const xml =
      '<worksheet><sheetData><row r="1"><c r="A1" cm="1" s="2">' +
      '<f t="array" ref="A1">[1]Sheet1!A1</f><v>5</v></c></row></sheetData></worksheet>'
    expect(transformSheetForLinkBreak(xml, [1])).toContain('<c r="A1" s="2"><v>5</v></c>')
  })

  it('renumbers [3] to [2] when [1] is broken, cell formulas included', () => {
    const xml =
      '<worksheet><sheetData><row r="1"><c r="A1"><f>[2]Sheet1!A1</f><v>1</v></c>' +
      '<c r="B1"><f>[3]Sheet1!A1</f><v>2</v></c></row></sheetData></worksheet>'
    const out = transformSheetForLinkBreak(xml, [1])
    expect(out).toContain('<f>[1]Sheet1!A1</f>')
    expect(out).toContain('<f>[2]Sheet1!A1</f>')
    expect(out).not.toContain('[3]Sheet1')
  })

  it('renumbers conditional-format and data-validation formulas too', () => {
    const xml =
      '<worksheet><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1">' +
      '<formula>[2]Sheet1!$A$1&gt;5</formula></cfRule></conditionalFormatting>' +
      '<dataValidations count="1"><dataValidation type="list">' +
      '<formula1>[3]Sheet1!$A$1:$A$3</formula1></dataValidation></dataValidations></worksheet>'
    const out = transformSheetForLinkBreak(xml, [1])
    expect(out).toContain('<formula>[1]Sheet1!$A$1&gt;5</formula>')
    expect(out).toContain('<formula1>[2]Sheet1!$A$1:$A$3</formula1>')
  })

  it('refuses a break a conditional format references instead of leaving it dangling', () => {
    const xml =
      '<worksheet><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1">' +
      '<formula>[1]Sheet1!$A$1&gt;5</formula></cfRule></conditionalFormatting></worksheet>'
    expect(() => transformSheetForLinkBreak(xml, [1])).toThrow(/conditional-formatting rule/)
  })

  it('refuses a break a data validation references', () => {
    const xml =
      '<worksheet><dataValidations count="1"><dataValidation type="list">' +
      '<formula1>[1]Sheet1!$A$1</formula1></dataValidation></dataValidations></worksheet>'
    expect(() => transformSheetForLinkBreak(xml, [1])).toThrow(/data-validation rule/)
  })

  it('still refuses a shared-formula anchor that references the broken link', () => {
    const xml =
      '<worksheet><sheetData><row r="1">' +
      '<c r="A1"><f t="shared" ref="A1:A2" si="0">[1]Sheet1!A1</f><v>1</v></c>' +
      '<c r="A2"><f t="shared" si="0"/><v>2</v></c></row></sheetData></worksheet>'
    expect(() => transformSheetForLinkBreak(xml, [1])).toThrow(/shared formula/)
  })
})

describe('transformWorkbookForLinkBreak', () => {
  it('renumbers surviving defined names', () => {
    const xml =
      '<workbook><definedNames><definedName name="Live">[3]Sheet1!$A$1</definedName>' +
      '</definedNames></workbook>'
    expect(transformWorkbookForLinkBreak(xml, [1])).toContain('[2]Sheet1!$A$1')
  })

  it('refuses a defined name that references the broken link', () => {
    const xml =
      '<workbook><definedNames><definedName name="Dead">[1]Sheet1!$A$1</definedName>' +
      '</definedNames></workbook>'
    expect(() => transformWorkbookForLinkBreak(xml, [1])).toThrow(/defined name/)
  })
})

describe('transformChartForLinkBreak', () => {
  it('renumbers surviving chart ranges', () => {
    const xml = '<c:chartSpace><c:f>[3]Sheet1!$A$1:$A$3</c:f></c:chartSpace>'
    expect(transformChartForLinkBreak(xml, [1])).toContain('<c:f>[2]Sheet1!$A$1:$A$3</c:f>')
  })

  it('refuses a chart range that references the broken link', () => {
    const xml = '<c:chartSpace><c:f>[1]Sheet1!$A$1:$A$3</c:f></c:chartSpace>'
    expect(() => transformChartForLinkBreak(xml, [1])).toThrow(/chart range/)
  })
})

describe('planExternalLinkBreak', () => {
  it('breaks one link across cells, defined names, formats and charts', async () => {
    const plan = await planExternalLinkBreak(
      await createBufferEntrySource(await buildWorkbook()),
      [1],
    )
    const workbook = plan.replaced.get('xl/workbook.xml')!
    expect(workbook).not.toContain('rId4')
    expect(workbook).toContain('<externalReference r:id="rId5"/>')
    // the defined name pointed at [3]; with [1] gone it is [2]
    expect(workbook).toContain('<definedName name="Renewal">[2]Sheet1!$A$1</definedName>')
    expect(plan.replaced.get('xl/_rels/workbook.xml.rels')).not.toContain('externalLink1.xml')
    expect(plan.replaced.get('[Content_Types].xml')).not.toContain(
      '/xl/externalLinks/externalLink1.xml',
    )
    expect(plan.removedEntries).toContain('xl/externalLinks/externalLink1.xml')

    const sheet = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(sheet).toContain('<c r="A1" s="2"/>')
    expect(sheet).toContain('<f>[1]Sheet1!A1*2</f>')
    expect(sheet).toContain('<f>[2]Sheet1!A1&amp;"!"</f>')
    expect(sheet).toContain('<formula>[1]Sheet1!A1&gt;5</formula>')

    const chart = plan.replaced.get('xl/charts/chart1.xml')!
    expect(chart).toContain('<c:f>[2]Sheet1!$A$1:$A$3</c:f>')
  })

  it('refuses the whole break when a defined name references the broken link', async () => {
    const source = await buildWorkbook()
    // rewrite the name to point at the link being broken
    const zip = await JSZip.loadAsync(source)
    const workbook = (await zip.file('xl/workbook.xml')!.async('string')).replace(
      '[3]Sheet1!$A$1',
      '[1]Sheet1!$A$1',
    )
    zip.file('xl/workbook.xml', workbook)
    const broken = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    await expect(planExternalLinkBreak(await createBufferEntrySource(broken), [1])).rejects.toThrow(
      /defined name/,
    )
  })
})
