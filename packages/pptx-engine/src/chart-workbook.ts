/**
 * The workbook embedded beside a chart part, which PowerPoint and Word open for
 * "Edit Data": one Sheet1 with the categories in column A and one column per
 * series from B, the cells the chart's c:f references point at.
 *
 *   A1        | B1 (ser 0 name) | C1 (ser 1 name) ...
 *   A2 (cat0) | B2 (val 0,0)    | C2 (val 1,0)   ...
 */
import { crc32 } from './crc32'
import { escapeXmlText } from './xml-utils'

export const CHART_WORKBOOK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
export const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// Series data columns start at B (column A holds the categories):
// 0 → B … 24 → Z, 25 → AA, 26 → AB, … (plain charCode arithmetic breaks past Z).
export const colLetter = (i: number): string => {
  let n = i + 2 // 1-based column number: A = 1, B = 2
  let label = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    label = String.fromCharCode(65 + rem) + label
    n = Math.floor((n - 1) / 26)
  }
  return label
}

export interface ChartWorkbookData {
  categories: string[]
  series: Array<{ name: string; values: (number | null | undefined)[] }>
  /** column A as numbers instead of category names: a scatter chart's x values */
  xValues?: number[]
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

/** The workbook's parts, by path inside the xlsx package. */
export function chartWorkbookParts({
  categories,
  series,
  xValues,
}: ChartWorkbookData): Record<string, string> {
  const strings: string[] = []
  const si = (s: string): number => {
    const idx = strings.indexOf(s)
    if (idx !== -1) return idx
    strings.push(s)
    return strings.length - 1
  }

  const headerCells = [`<c r="A1" t="s"><v>${si('')}</v></c>`]
  series.forEach((ser, j) => {
    headerCells.push(`<c r="${colLetter(j)}1" t="s"><v>${si(ser.name)}</v></c>`)
  })
  const dataRows = categories.map((category, i) => {
    const rowNum = i + 2
    const x = xValues?.[i]
    const cells = [
      x !== undefined
        ? `<c r="A${rowNum}"><v>${x}</v></c>`
        : `<c r="A${rowNum}" t="s"><v>${si(category)}</v></c>`,
    ]
    series.forEach((ser, j) => {
      const val = ser.values[i]
      if (val != null && Number.isFinite(val)) {
        cells.push(`<c r="${colLetter(j)}${rowNum}"><v>${val}</v></c>`)
      }
    })
    return `<row r="${rowNum}">${cells.join('')}</row>`
  })

  return {
    '[Content_Types].xml':
      XML_DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '</Types>',
    '_rels/.rels':
      XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
    'xl/workbook.xml':
      XML_DECL +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
      '</workbook>',
    'xl/_rels/workbook.xml.rels':
      XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
      '</Relationships>',
    'xl/worksheets/sheet1.xml':
      XML_DECL +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      `<sheetData><row r="1">${headerCells.join('')}</row>${dataRows.join('')}</sheetData>` +
      '</worksheet>',
    'xl/sharedStrings.xml':
      XML_DECL +
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
      strings.map((s) => `<si><t xml:space="preserve">${escapeXmlText(s)}</t></si>`).join('') +
      '</sst>',
  }
}

/**
 * A zip of `files`, stored without compression: the chart engines insert charts
 * synchronously, and a chart's workbook is a few hundred bytes.
 */
export function storedZip(files: Record<string, string>): Uint8Array {
  const utf8 = new TextEncoder()
  const entries = Object.entries(files).map(([path, text]) => {
    const name = utf8.encode(path)
    const data = utf8.encode(text)
    return { name, data, crc: crc32(data) }
  })
  const localSize = entries.reduce((n, e) => n + 30 + e.name.length + e.data.length, 0)
  const centralSize = entries.reduce((n, e) => n + 46 + e.name.length, 0)
  const out = new Uint8Array(localSize + centralSize + 22)
  const view = new DataView(out.buffer)
  let local = 0
  let central = localSize
  for (const { name, data, crc } of entries) {
    view.setUint32(local, 0x04034b50, true)
    view.setUint16(local + 4, 10, true)
    view.setUint16(local + 6, 0x0800, true)
    view.setUint16(local + 12, 0x21, true) // 1980-01-01, the DOS epoch
    view.setUint32(local + 14, crc, true)
    view.setUint32(local + 18, data.length, true)
    view.setUint32(local + 22, data.length, true)
    view.setUint16(local + 26, name.length, true)
    out.set(name, local + 30)
    out.set(data, local + 30 + name.length)

    view.setUint32(central, 0x02014b50, true)
    view.setUint16(central + 4, 20, true)
    view.setUint16(central + 6, 10, true)
    view.setUint16(central + 8, 0x0800, true)
    view.setUint16(central + 14, 0x21, true)
    view.setUint32(central + 16, crc, true)
    view.setUint32(central + 20, data.length, true)
    view.setUint32(central + 24, data.length, true)
    view.setUint16(central + 28, name.length, true)
    view.setUint32(central + 42, local, true)
    out.set(name, central + 46)

    local += 30 + name.length + data.length
    central += 46 + name.length
  }
  view.setUint32(central, 0x06054b50, true)
  view.setUint16(central + 8, entries.length, true)
  view.setUint16(central + 10, entries.length, true)
  view.setUint32(central + 12, centralSize, true)
  view.setUint32(central + 16, localSize, true)
  return out
}

export function chartWorkbookXlsx(data: ChartWorkbookData): Uint8Array {
  return storedZip(chartWorkbookParts(data))
}
