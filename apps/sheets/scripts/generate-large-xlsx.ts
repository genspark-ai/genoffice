import { writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'

import JSZip from 'jszip'

/// Synthetic value-mode workbook with the requested number of stored cells
/// (10 columns: numbers, shared strings and styled cells, no formulas) for
/// the full-load benchmark. An optional bounding-box column count plants one
/// extra cell in that column on the last row (sparse box, dense data).
///   tsx scripts/generate-large-xlsx.ts <storedCells> <out.xlsx> [boxColumns]
const COLUMNS = 10
const SHARED = 200

async function main(): Promise<void> {
  const cells = Number(process.argv[2])
  const out = process.argv[3]
  const boxColumns = Math.max(COLUMNS, Number(process.argv[4] ?? COLUMNS))
  if (!Number.isInteger(cells) || cells <= 0 || !out || !Number.isInteger(boxColumns)) {
    throw new Error('usage: generate-large-xlsx.ts <storedCells> <out.xlsx> [boxColumns]')
  }
  const rows = Math.ceil(cells / COLUMNS)
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<dimension ref="A1:${columnLetter(boxColumns - 1)}${rows}"/>`,
    '<sheetData>',
  ]
  let written = 0
  for (let row = 1; row <= rows && written < cells; row += 1) {
    let line = `<row r="${row}">`
    if (row === rows && boxColumns > COLUMNS) {
      line += `<c r="${columnLetter(boxColumns - 1)}${row}"><v>1</v></c>`
      written += 1
    }
    for (let column = 0; column < COLUMNS && written < cells; column += 1) {
      const ref = `${columnLetter(column)}${row}`
      switch (column % 4) {
        case 0:
          line += `<c r="${ref}"><v>${row * 7 + column}</v></c>`
          break
        case 1:
          line += `<c r="${ref}" t="s"><v>${(row + column) % SHARED}</v></c>`
          break
        case 2:
          line += `<c r="${ref}" s="1"><v>${(row * 1.37 + column).toFixed(2)}</v></c>`
          break
        default:
          line += `<c r="${ref}" s="2"><v>${45000 + (row % 3650)}</v></c>`
      }
      written += 1
    }
    parts.push(line + '</row>')
  }
  parts.push('</sheetData></worksheet>')
  const sharedStrings =
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${SHARED}" uniqueCount="${SHARED}">` +
    Array.from({ length: SHARED }, (_, i) => `<si><t>Item ${i} lorem ipsum</t></si>`).join('') +
    '</sst>'
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="164" fontId="1" fillId="2" borderId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs></styleSheet>',
  )
  zip.file('xl/sharedStrings.xml', sharedStrings)
  zip.file('xl/worksheets/sheet1.xml', parts.join('\n'))
  const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  await writeFile(resolve(out), buffer)
  process.stdout.write(`${out}: ${written} stored cells, ${rows} rows, ${buffer.length} bytes\n`)
}

function columnLetter(index: number): string {
  let value = index + 1
  let letters = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    value = Math.floor((value - 1) / 26)
  }
  return letters
}

void main()
