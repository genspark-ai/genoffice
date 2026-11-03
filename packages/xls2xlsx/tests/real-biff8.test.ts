import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { convertXlsToXlsx, openXlsContainer, parseXlsWorkbook } from '../src/index'

/// A genuine BIFF8 workbook: a small .xlsx written by this repo's test builder
/// was converted with LibreOffice (soffice --convert-to xls). It carries the
/// layouts the synthetic byte builder used to mirror by hand: 4-byte SST
/// counts, a rich shared string with cRun before the characters, FormulaValue
/// results (number/string/bool/error), a zero-based Ref8 merge and f1904=1.
const FIXTURE = new URL('./fixtures/real-biff8.xls', import.meta.url)
const bytes = new Uint8Array(readFileSync(FIXTURE))

async function sheetXml(name = 'xl/worksheets/sheet1.xml'): Promise<string> {
  const out = await convertXlsToXlsx(bytes)
  const zip = await JSZip.loadAsync(out.bytes)
  return zip.file(name)!.async('string')
}

function cellValue(xml: string, ref: string): string | undefined {
  return new RegExp('<c r="' + ref + '"[^>]*><v>([^<]*)</v></c>').exec(xml)?.[1]
}

describe('a real LibreOffice-authored BIFF8 workbook', () => {
  it('parses the values, rich text and 1904 epoch from the SST/FormulaValue/Date1904 records', () => {
    const workbook = parseXlsWorkbook(openXlsContainer(bytes).workbook)
    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
    expect(workbook.date1904).toBe(true)
    const sheet = workbook.sheets[0]!
    const cell = (row: number, col: number) => sheet.rows.get(row)?.get(col)
    expect(cell(0, 0)).toEqual({ row: 0, col: 0, type: 'string', value: 'Name' })
    expect(cell(0, 1)).toEqual({ row: 0, col: 1, type: 'string', value: '中文标签' })
    expect(cell(0, 2)).toEqual({ row: 0, col: 2, type: 'string', value: 'BoldPlain' })
    expect(cell(0, 3)).toEqual({ row: 0, col: 3, type: 'string', value: 'Value' })
    expect(cell(2, 0)).toEqual({ row: 2, col: 0, type: 'number', value: 42.5 })
    // formula cached values: number, string, boolean, error (FormulaValue)
    expect(cell(1, 0)).toEqual({ row: 1, col: 0, type: 'number', value: 2 })
    expect(cell(1, 1)).toEqual({ row: 1, col: 1, type: 'string', value: 'text' })
    expect(cell(1, 2)).toEqual({ row: 1, col: 2, type: 'boolean', value: true })
    expect(cell(1, 3)).toEqual({ row: 1, col: 3, type: 'error', value: '#DIV/0!' })
    // Ref8 is zero-based: A6:C7 is rows 5..6, columns 0..2
    expect(sheet.merges).toEqual([{ firstRow: 5, lastRow: 6, firstCol: 0, lastCol: 2 }])
    expect(sheet.colWidths.get(0)).toBeCloseTo(15.24, 1)
  })

  it('writes the values, the merge and the 1904 workbook property into the xlsx', async () => {
    const xml = await sheetXml()
    expect(cellValue(xml, 'A1')).toBe('0') // shared string index for "Name"
    expect(cellValue(xml, 'B1')).toBe('1') // the CJK label cell
    expect(cellValue(xml, 'C1')).toBe('2') // "BoldPlain" (rich string flattened)
    expect(cellValue(xml, 'A2')).toBe('2')
    expect(cellValue(xml, 'A3')).toBe('42.5')
    expect(xml).toContain('<c r="B2" t="s"><v>4</v></c>')
    expect(xml).toContain('<c r="C2" t="b"><v>1</v></c>')
    expect(xml).toContain('<c r="D2" t="e"><v>#DIV/0!</v></c>')
    expect(xml).toContain('<mergeCell ref="A6:C7"/>')
    expect(xml).not.toContain('NaN')
    const workbookXml = await sheetXml('xl/workbook.xml')
    expect(workbookXml).toContain('<workbookPr date1904="1"/>')
    const out = await convertXlsToXlsx(bytes)
    expect(out.warnings.join('\n')).not.toContain('date system')
  })
})
