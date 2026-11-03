import JSZip from 'jszip'
import type { XlsWorkbook } from './workbook'

/**
 * The minimal OOXML writer the conversion lands in: workbook, sheets with
 * their values, shared strings, merges, column widths — and the stub styles
 * part Excel requires before it will open anything. Formula cells carry
 * their cached value only (the RPN translation is a later stage), which the
 * `<c>` element expresses by simply not writing an `<f>`.
 */

const COLUMN_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

export function columnName(index: number): string {
  let name = ''
  let n = index
  while (n >= 0) {
    name = COLUMN_LETTERS[n % 26] + name
    n = Math.floor(n / 26) - 1
  }
  return name
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

export async function writeXlsxWorkbook(workbook: XlsWorkbook): Promise<Uint8Array> {
  const zip = new JSZip()
  const sharedStrings: string[] = []
  const stringIndex = new Map<string, number>()

  const intern = (text: string): number => {
    const existing = stringIndex.get(text)
    if (existing !== undefined) return existing
    const index = sharedStrings.length
    sharedStrings.push(text)
    stringIndex.set(text, index)
    return index
  }

  const sheetXmls = workbook.sheets.map((sheet) => {
    const rowsXml: string[] = []
    const rowNumbers = [...sheet.rows.keys()].sort((a, b) => a - b)
    for (const rowNumber of rowNumbers) {
      const cells = sheet.rows.get(rowNumber)!
      const cellXml: string[] = []
      const cols = [...cells.keys()].sort((a, b) => a - b)
      for (const col of cols) {
        const cell = cells.get(col)!
        const ref = `${columnName(col)}${rowNumber + 1}`
        if (cell.type === 'blank') continue
        if (cell.type === 'string') {
          if (cell.value === '') continue
          cellXml.push(`<c r="${ref}" t="s"><v>${intern(cell.value as string)}</v></c>`)
        } else if (cell.type === 'number') {
          cellXml.push(`<c r="${ref}"><v>${cell.value}</v></c>`)
        } else if (cell.type === 'boolean') {
          cellXml.push(`<c r="${ref}" t="b"><v>${cell.value ? 1 : 0}</v></c>`)
        } else {
          cellXml.push(`<c r="${ref}" t="e"><v>${escapeXml(String(cell.value))}</v></c>`)
        }
      }
      if (cellXml.length === 0) continue
      rowsXml.push(`<row r="${rowNumber + 1}">${cellXml.join('')}</row>`)
    }
    const dimension = sheet.dimension
    const dimAttr =
      dimension &&
      dimension.lastRow >= dimension.firstRow &&
      dimension.lastCol >= dimension.firstCol
        ? `<dimension ref="${columnName(dimension.firstCol)}${dimension.firstRow + 1}:${columnName(
            dimension.lastCol,
          )}${dimension.lastRow + 1}"/>`
        : ''
    const colsXml =
      sheet.colWidths.size > 0
        ? `<cols>${[...sheet.colWidths.entries()]
            .map(
              ([col, width]) =>
                `<col min="${col + 1}" max="${col + 1}" width="${width}" customWidth="1"/>`,
            )
            .join('')}</cols>`
        : ''
    const mergesXml =
      sheet.merges.length > 0
        ? `<mergeCells count="${sheet.merges.length}">${sheet.merges
            .map(
              (m) =>
                `<mergeCell ref="${columnName(m.firstCol)}${m.firstRow + 1}:${columnName(
                  m.lastCol,
                )}${m.lastRow + 1}"/>`,
            )
            .join('')}</mergeCells>`
        : ''
    return {
      name: sheet.name,
      visible: sheet.visible,
      xml:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        dimAttr +
        colsXml +
        `<sheetData>${rowsXml.join('')}</sheetData>` +
        mergesXml +
        '</worksheet>',
    }
  })

  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheetXmls
        .map(
          (_s, i) =>
            `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join('') +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  // The 1904 date system is a workbook property in OOXML; without it every
  // date serial from a Mac-authored .xls shifts by four years.
  const workbookPr = workbook.date1904 ? '<workbookPr date1904="1"/>' : ''
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      workbookPr +
      '<bookViews><workbookView/></bookViews>' +
      '<sheets>' +
      sheetXmls
        .map(
          (s, i) =>
            `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"${
              s.visible ? '' : ' state="hidden"'
            }/>`,
        )
        .join('') +
      '</sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheetXmls
        .map(
          (_s, i) =>
            `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
        )
        .join('') +
      `<Relationship Id="rId${sheetXmls.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
      `<Relationship Id="rId${sheetXmls.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      '</Relationships>',
  )
  sheetXmls.forEach((s, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, s.xml))
  zip.file(
    'xl/sharedStrings.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">` +
      sharedStrings.map((text) => `<si><t>${escapeXml(text)}</t></si>`).join('') +
      '</sst>',
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
      '</styleSheet>',
  )

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
