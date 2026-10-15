import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { clipSheetXml, xlsxToText } from '../src/xlsx'

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

function workbook(sheets: string[]): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    `<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
      sheets
        .map((_, i) => `<sheet name="S${i + 1}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
        .join('') +
      '</sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets
        .map(
          (_, i) =>
            `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
        )
        .join('') +
      '</Relationships>',
  )
  sheets.forEach((rows, i) => {
    zip.file(
      `xl/worksheets/sheet${i + 1}.xml`,
      `<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData><pageMargins left="0.7"/></worksheet>`,
    )
  })
  return zip.generateAsync({ type: 'uint8array' })
}

const rows = (n: number, tag: string): string =>
  Array.from(
    { length: n },
    (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${tag}${i + 1}</t></is></c></row>`,
  ).join('')

describe('xlsxToText maxChars', () => {
  it('keeps every row without a budget', async () => {
    const text = await xlsxToText(await workbook([rows(50_000, 'row')]))
    expect(text).toContain('row50000')
  })

  it('stops at the budget and skips the sheets after it', async () => {
    const text = await xlsxToText(await workbook([rows(50_000, 'row'), rows(5, 'other')]), {
      maxChars: 2_000,
    })
    expect(text).toContain('row1\n')
    expect(text).not.toContain('row50000')
    expect(text).not.toContain('# S2')
    expect(text.length).toBeLessThan(2_200)
  })

  it('keeps the leading rows of a prefixed sheet under a budget', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      `<x:workbook xmlns:x="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheets><x:sheet name="P" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>`,
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    )
    const body = Array.from(
      { length: 50_000 },
      (_, i) =>
        `<x:row r="${i + 1}"><x:c r="A${i + 1}" t="inlineStr"><x:is><x:t>p${i + 1}</x:t></x:is></x:c></x:row>`,
    ).join('')
    zip.file(
      'xl/worksheets/sheet1.xml',
      `<x:worksheet xmlns:x="${NS}"><x:sheetData>${body}</x:sheetData></x:worksheet>`,
    )
    const text = await xlsxToText(await zip.generateAsync({ type: 'uint8array' }), {
      maxChars: 2_000,
    })
    expect(text).toContain('p1\n')
    expect(text).toContain('p100\n')
    expect(text).not.toContain('p50000')
  })

  it('reads a later sheet when the first one is small', async () => {
    const text = await xlsxToText(await workbook([rows(3, 'a'), rows(3, 'b')]), { maxChars: 2_000 })
    expect(text).toContain('# S2')
    expect(text).toContain('b3')
  })
})

describe('clipSheetXml', () => {
  it('leaves short XML alone', () => {
    const xml = `<worksheet><sheetData>${rows(3, 'r')}</sheetData></worksheet>`
    expect(clipSheetXml(xml, 10)).toBe(xml)
  })

  it('cuts on a row boundary and closes the document', () => {
    const xml = `<worksheet><sheetData>${rows(100_000, 'r')}</sheetData><cols/></worksheet>`
    const clipped = clipSheetXml(xml, 100)
    expect(clipped.length).toBeLessThan(xml.length)
    expect(clipped.endsWith('</row></sheetData></worksheet>')).toBe(true)
    expect(clipped).not.toContain('<cols/>')
  })

  it('closes a truncated prefix at its last complete row', () => {
    const xml = `<worksheet><sheetData>${rows(5, 'r')}<row r="6"><c r="A6" t="inlineStr"><is><t>r`
    const clipped = clipSheetXml(xml, 1_000_000, true)
    expect(clipped.endsWith('r5</t></is></c></row></sheetData></worksheet>')).toBe(true)
  })

  it('yields an empty sheet when a truncated prefix holds no complete row', () => {
    expect(clipSheetXml('<worksheet><sheetData><row r="1"><c r="A1"', 10, true)).toBe(
      '<worksheet><sheetData/></worksheet>',
    )
  })

  it('handles namespace-prefixed worksheet markup', () => {
    const body = Array.from(
      { length: 100_000 },
      (_, i) =>
        `<x:row r="${i + 1}"><x:c r="A${i + 1}" t="inlineStr"><x:is><x:t>r${i + 1}</x:t></x:is></x:c></x:row>`,
    ).join('')
    const xml = `<x:worksheet xmlns:x="${NS}"><x:sheetData>${body}</x:sheetData></x:worksheet>`
    const clipped = clipSheetXml(xml, 100, true)
    expect(clipped.length).toBeLessThan(xml.length)
    expect(clipped.endsWith('</x:row></x:sheetData></x:worksheet>')).toBe(true)
  })

  it('does not cut before the first row', () => {
    const head = 'x'.repeat(300 * 1024)
    const xml = `<worksheet><!--${head}--><sheetData>${rows(2, 'r')}</sheetData></worksheet>`
    expect(clipSheetXml(xml, 1)).toBe(xml)
  })
})
