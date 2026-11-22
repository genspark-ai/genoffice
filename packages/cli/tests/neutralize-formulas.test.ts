import { readFileSync, writeFileSync } from 'node:fs'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

/** A minimal workbook whose A-column cells are text (inlineStr), not formulas. */
async function writeWorkbook(path: string, texts: string[]): Promise<void> {
  const rows = texts
    .map(
      (t, i) =>
        `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t xml:space="preserve">${t
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')}</t></is></c></row>`,
    )
    .join('')
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="data" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`,
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

// the round-2 misses: the function call sits anywhere behind the prefix, not
// right after it; the accepted ordinary data keeps its bytes on both paths
const TEXTS = [
  '+1+WEBSERVICE(A1)',
  '-1+HYPERLINK(A1)',
  '-(HYPERLINK(A1))',
  '+A1&WEBSERVICE(A1)',
  '-$1,234.50',
  '- item',
  '@handle',
  '+86 138',
]

describe('convert --to csv --neutralize-formulas', () => {
  it.skipIf(!xlsxSidecarPath())(
    'prefixes the executing shapes and leaves ordinary data untouched (default off)',
    async () => {
      const dir = tempDir()
      const xlsx = `${dir}/guard.xlsx`
      await writeWorkbook(xlsx, TEXTS)

      const plain = await run([
        'convert',
        xlsx,
        '--to',
        'csv',
        '--out',
        `${dir}/plain.csv`,
        '--json',
      ])
      expect(plain.code).toBe(0)
      const plainCsv = readFileSync(`${dir}/plain.csv`, 'utf8')
      for (const t of TEXTS) expect(plainCsv).toContain(t)

      const guarded = await run([
        'convert',
        xlsx,
        '--to',
        'csv',
        '--out',
        `${dir}/safe.csv`,
        '--neutralize-formulas',
        '--json',
      ])
      expect(guarded.code).toBe(0)
      const safeCsv = readFileSync(`${dir}/safe.csv`, 'utf8')
      // every executing shape gains the text-marker apostrophe
      expect(safeCsv).toContain("'+1+WEBSERVICE(A1)")
      expect(safeCsv).toContain("'-1+HYPERLINK(A1)")
      expect(safeCsv).toContain("'-(HYPERLINK(A1))")
      expect(safeCsv).toContain("'+A1&WEBSERVICE(A1)")
      // ordinary data keeps its bytes
      expect(safeCsv).toContain('-$1,234.50')
      expect(safeCsv).toContain('- item')
      expect(safeCsv).toContain('@handle')
      expect(safeCsv).toContain('+86 138')
    },
  )
})
