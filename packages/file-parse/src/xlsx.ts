import JSZip from 'jszip'
import { assertZipWithinLimits } from '@genoffice/docx-engine'
import { resolveTarget } from './opc'
import { XMLParser } from 'fast-xml-parser'
import {
  builtinDateFormat,
  classifyFormatCode,
  formatSerial,
  type DateFormatParts,
} from './xlsx-dates'

function stripNamespacePrefix(name: string): string {
  const separator = name.indexOf(':')
  return separator < 0 ? name : name.slice(separator + 1)
}

// Text fidelity: no trim (xml:space="preserve" runs carry the spaces between words),
// no numeric coercion of tag values (otherwise <t>02139</t> becomes a number and loses characters)
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  removeNSPrefix: false,
  transformTagName: stripNamespacePrefix,
  // trimValues also governs attributes; nothing read here (r:id, Target, cell ref, sheet
  // name) carries meaningful edge whitespace, and an untrimmed Target builds the wrong zip
  // path, which silently drops the whole sheet
  attributeValueProcessor: (_name, value) => value.trim(),
})

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? value : [value]
}

function relationshipId(node: Record<string, unknown>): string {
  const qualified = Object.entries(node).find(([key]) => /^@_[^:]+:id$/.test(key))
  return String(qualified?.[1] ?? node['@_id'] ?? '')
}

/** unwrap a fast-xml-parser text node (plain value, or { '#text': ... } when it carried attributes) */
function textOf(node: unknown): string {
  if (node == null) return ''
  if (typeof node === 'object' && !Array.isArray(node)) {
    return textOf((node as Record<string, unknown>)['#text'])
  }
  return String(node)
}

/** shared string entry: <si><t>…</t></si> or rich-text runs <si><r><t>…</t></r>…</si> */
function sharedStringText(si: Record<string, unknown>): string {
  let out = ''
  for (const key of Object.keys(si)) {
    if (key === 't') {
      for (const node of asArray(si[key] as unknown)) out += textOf(node)
    } else if (key === 'r') {
      for (const run of asArray(si[key] as unknown)) {
        const r = run as Record<string, unknown>
        for (const node of asArray(r['t'] as unknown)) out += textOf(node)
      }
    }
  }
  return out
}

/** "BC12" → zero-based column index 54 (cell refs are case-insensitive per ECMA-376) */
function columnIndex(cellRef: string): number {
  let index = 0
  let letters = 0
  for (const ch of cellRef.toUpperCase()) {
    if (ch < 'A' || ch > 'Z') break
    letters += 1
    index = index * 26 + (ch.charCodeAt(0) - 64)
  }
  if (letters === 0) return -1
  return index - 1
}

interface Cell {
  '@_r'?: string
  '@_t'?: string
  '@_s'?: string
  v?: unknown
  f?: unknown
  is?: Record<string, unknown>
}

type SharedFormulas = Map<string, string>

function sharedIndex(cell: Cell): string | undefined {
  const si =
    typeof cell.f === 'object' && cell.f !== null
      ? (cell.f as Record<string, unknown>)['@_si']
      : undefined
  return si === undefined ? undefined : String(si)
}

/** the master cell may well carry a cached value, so register it before the value check */
function recordSharedFormula(cell: Cell, sharedFormulas: SharedFormulas): void {
  const si = sharedIndex(cell)
  const own = textOf(cell.f).trim()
  if (si !== undefined && own) sharedFormulas.set(si, own)
}

/** openpyxl/ExcelJS write formulas with no cached <v>; the formula text is all there is until Excel recalculates */
function formulaText(cell: Cell, sharedFormulas: SharedFormulas): string {
  const si = sharedIndex(cell)
  const text = textOf(cell.f).trim() || (si !== undefined ? (sharedFormulas.get(si) ?? '') : '')
  return text ? `=${text}` : ''
}

/** cellXfs index → date/time parts, for the xf entries whose numFmt renders a calendar value */
type DateStyles = Map<number, DateFormatParts>

/**
 * A date cell is stored as a plain serial; without its style the model sees 45292 where the
 * sheet shows 1-Jan-24. Resolve each cellXfs entry's numFmt once per workbook.
 */
async function loadDateStyles(zip: JSZip): Promise<DateStyles> {
  const styles: DateStyles = new Map()
  const xml = await zipText(zip, 'xl/styles.xml')
  if (!xml) return styles
  const sheet = (parser.parse(xml) as Record<string, any>).styleSheet
  const custom = new Map<number, DateFormatParts | null>()
  for (const fmt of asArray(sheet?.numFmts?.numFmt) as Array<Record<string, unknown>>) {
    const id = Number(fmt['@_numFmtId'])
    if (Number.isInteger(id)) custom.set(id, classifyFormatCode(String(fmt['@_formatCode'] ?? '')))
  }
  const xfs = asArray(sheet?.cellXfs?.xf) as Array<Record<string, unknown>>
  xfs.forEach((xf, index) => {
    const id = Number(xf['@_numFmtId'] ?? 0)
    const parts = custom.has(id) ? custom.get(id) : builtinDateFormat(id)
    if (parts) styles.set(index, parts)
  })
  return styles
}

function isDate1904(workbook: Record<string, any>): boolean {
  const flag = String(workbook.workbook?.workbookPr?.['@_date1904'] ?? '')
    .trim()
    .toLowerCase()
  return flag === '1' || flag === 'true'
}

function cellText(
  cell: Cell,
  shared: string[],
  dates: DateStyles,
  date1904: boolean,
  sharedFormulas: SharedFormulas,
): string {
  const type = cell['@_t'] ?? ''
  if (type === 'inlineStr') return cell.is ? sharedStringText(cell.is) : ''
  if (cell.f !== undefined) {
    recordSharedFormula(cell, sharedFormulas)
    if (textOf(cell.v).trim() === '') return formulaText(cell, sharedFormulas)
  }
  // <v> is ST_Xstring so it reaches the caller verbatim; the two reads that need it as a
  // scalar handle their own whitespace (Number tolerates it, the boolean compare strips it)
  const value = textOf(cell.v)
  if ((type === '' || type === 'n') && cell['@_s'] !== undefined) {
    const parts = dates.get(Number(cell['@_s']))
    if (parts && value.trim() !== '') {
      const rendered = formatSerial(Number(value), parts, date1904)
      if (rendered !== null) return rendered
    }
  }
  if (type === 's') {
    // An empty or missing <v> must stay empty: Number('') is 0 and would
    // otherwise leak shared[0] into the cell. Out-of-range or non-numeric
    // indexes also degrade to empty rather than corrupting the row.
    const trimmed = value.trim()
    if (trimmed === '') return ''
    const index = Number(trimmed)
    if (!Number.isInteger(index) || index < 0 || index >= shared.length) return ''
    return shared[index] ?? ''
  }
  if (type === 'b') return value.trim() === '1' ? 'TRUE' : 'FALSE'
  return value
}

async function zipText(zip: JSZip, path: string): Promise<string | undefined> {
  const file = zip.file(path)
  return file ? file.async('text') : undefined
}

/** Malformed refs (e.g. XXXXXX99) would otherwise grow the cells array by
 *  millions via the padding loop below. */
export const MAX_XLSX_COLS = 16_384

/** extract sheet text from an xlsx: one "# SheetName" section per sheet, cells joined with " | " */
export async function xlsxToText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  assertZipWithinLimits(zip)
  const workbookXml = await zipText(zip, 'xl/workbook.xml')
  if (!workbookXml) throw new Error('Invalid xlsx: missing xl/workbook.xml')

  // Sheet order and names come from workbook.xml; r:id maps to the actual sheet path via the workbook rels
  const workbook = parser.parse(workbookXml) as Record<string, any>
  const sheets = asArray(workbook.workbook?.sheets?.sheet) as Array<Record<string, unknown>>
  const dateStyles = await loadDateStyles(zip)
  const date1904 = isDate1904(workbook)

  const relsXml = await zipText(zip, 'xl/_rels/workbook.xml.rels')
  const relTargets = new Map<string, string>()
  if (relsXml) {
    const rels = parser.parse(relsXml) as Record<string, any>
    for (const rel of asArray(rels.Relationships?.Relationship) as Array<Record<string, unknown>>) {
      const target = String(rel['@_Target'] ?? '')
      relTargets.set(String(rel['@_Id'] ?? ''), resolveTarget('xl/workbook.xml', target))
    }
  }

  const shared: string[] = []
  const sharedXml = await zipText(zip, 'xl/sharedStrings.xml')
  if (sharedXml) {
    const sst = parser.parse(sharedXml) as Record<string, any>
    for (const si of asArray(sst.sst?.si) as Array<Record<string, unknown>>) {
      shared.push(sharedStringText(si))
    }
  }

  const sections: string[] = []
  let sheetsWithData = 0
  let imageOnlySheets = 0
  for (const sheet of sheets) {
    const path = relTargets.get(relationshipId(sheet))
    const sheetXml = path ? await zipText(zip, path) : undefined
    if (!sheetXml) continue
    const worksheet = parser.parse(sheetXml) as Record<string, any>
    const lines: string[] = [`# ${String(sheet['@_name'] ?? '')}`]
    const rows = asArray(worksheet.worksheet?.sheetData?.row) as Array<Record<string, unknown>>
    let hasData = false
    const sharedFormulas: SharedFormulas = new Map()
    for (const row of rows) {
      const cells: string[] = []
      for (const cell of asArray(row.c as Cell | Cell[])) {
        const text = cellText(cell, shared, dateStyles, date1904, sharedFormulas)
        const ref = cell['@_r']
        // A malformed ref (no leading column letters) yields -1; append in
        // document order instead of writing cells[-1] which would drop text.
        // Clamp wild columns (e.g. XXXXXX99) to append: padding millions of
        // empty cells would OOM on a hostile file.
        const col = ref ? columnIndex(ref) : cells.length
        const target = col >= 0 && col < MAX_XLSX_COLS ? col : cells.length
        while (cells.length < target) cells.push('')
        cells[target] = text
        if (text.trim()) hasData = true
      }
      lines.push(cells.join(' | '))
    }
    if (hasData) sheetsWithData += 1
    else if (path && worksheet.worksheet?.drawing) {
      imageOnlySheets += 1
      const pictures = await countDrawingPictures(zip, path, worksheet.worksheet.drawing)
      lines.push(
        pictures > 0
          ? `[image-only sheet: ${pictures} image${pictures === 1 ? '' : 's'}, no cell data]`
          : '[image-only sheet: a drawing but no cell data]',
      )
    }
    sections.push(lines.join('\n'))
  }
  const body = sections.join('\n\n')
  if (sheetsWithData > 0 || imageOnlySheets === 0) return body
  const n = sections.length
  return `[No extractable text: none of the ${n} sheet${n === 1 ? '' : 's'} holds cell data; the content is in embedded images, which this extraction does not read.]\n\n${body}`
}

function countElements(value: unknown, name: string): number {
  if (Array.isArray(value)) return value.reduce((n, item) => n + countElements(item, name), 0)
  if (value === null || typeof value !== 'object') return 0
  return Object.entries(value).reduce((n, [key, child]) => {
    const matches = key === name ? (Array.isArray(child) ? child.length : 1) : 0
    return n + matches + countElements(child, name)
  }, 0)
}

/** follow the sheet's <drawing r:id> to its drawing part and count the pictures anchored there */
async function countDrawingPictures(
  zip: JSZip,
  sheetPath: string,
  drawing: Record<string, unknown>,
): Promise<number> {
  const relsPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels')
  const relsXml = await zipText(zip, relsPath)
  if (!relsXml) return 0
  const rels = parser.parse(relsXml) as Record<string, any>
  const wanted = relationshipId(drawing)
  for (const rel of asArray(rels.Relationships?.Relationship) as Array<Record<string, unknown>>) {
    if (String(rel['@_Id'] ?? '') !== wanted) continue
    const drawingXml = await zipText(zip, resolveTarget(sheetPath, String(rel['@_Target'] ?? '')))
    return drawingXml ? countElements(parser.parse(drawingXml), 'pic') : 0
  }
  return 0
}
