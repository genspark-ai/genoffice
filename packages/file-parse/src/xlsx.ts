import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'

// Text fidelity: no trim (xml:space="preserve" runs carry the spaces between words),
// no numeric coercion of tag values (otherwise <t>02139</t> becomes a number and loses characters)
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  // trimValues also governs attributes; nothing read here (r:id, Target, cell ref, sheet
  // name) carries meaningful edge whitespace, and an untrimmed Target builds the wrong zip
  // path, which silently drops the whole sheet
  attributeValueProcessor: (_name, value) => value.trim(),
})

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? value : [value]
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
  if (si['t'] !== undefined) return textOf(si['t'])
  return asArray(si['r'] as unknown)
    .map((run) => textOf((run as Record<string, unknown>)['t']))
    .join('')
}

/** "BC12" → zero-based column index 54 */
function columnIndex(cellRef: string): number {
  let index = 0
  for (const ch of cellRef) {
    if (ch < 'A' || ch > 'Z') break
    index = index * 26 + (ch.charCodeAt(0) - 64)
  }
  return index - 1
}

interface Cell {
  '@_r'?: string
  '@_t'?: string
  v?: unknown
  is?: Record<string, unknown>
}

function cellText(cell: Cell, shared: string[]): string {
  const type = cell['@_t'] ?? ''
  if (type === 'inlineStr') return cell.is ? sharedStringText(cell.is) : ''
  // <v> is ST_Xstring so it reaches the caller verbatim; the two reads that need it as a
  // scalar handle their own whitespace (Number tolerates it, the boolean compare strips it)
  const value = textOf(cell.v)
  if (type === 's') return shared[Number(value)] ?? ''
  if (type === 'b') return value.trim() === '1' ? 'TRUE' : 'FALSE'
  return value
}

async function zipText(zip: JSZip, path: string): Promise<string | undefined> {
  const file = zip.file(path)
  return file ? file.async('text') : undefined
}

/** extract sheet text from an xlsx: one "# SheetName" section per sheet, cells joined with " | " */
export async function xlsxToText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const workbookXml = await zipText(zip, 'xl/workbook.xml')
  if (!workbookXml) throw new Error('Invalid xlsx: missing xl/workbook.xml')

  // Sheet order and names come from workbook.xml; r:id maps to the actual sheet path via the workbook rels
  const workbook = parser.parse(workbookXml) as Record<string, any>
  const sheets = asArray(workbook.workbook?.sheets?.sheet) as Array<Record<string, unknown>>

  const relsXml = await zipText(zip, 'xl/_rels/workbook.xml.rels')
  const relTargets = new Map<string, string>()
  if (relsXml) {
    const rels = parser.parse(relsXml) as Record<string, any>
    for (const rel of asArray(rels.Relationships?.Relationship) as Array<Record<string, unknown>>) {
      const target = String(rel['@_Target'] ?? '')
      relTargets.set(
        String(rel['@_Id'] ?? ''),
        target.startsWith('/') ? target.slice(1) : `xl/${target}`,
      )
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
  for (const sheet of sheets) {
    const path = relTargets.get(String(sheet['@_r:id'] ?? ''))
    const sheetXml = path ? await zipText(zip, path) : undefined
    if (!sheetXml) continue
    const worksheet = parser.parse(sheetXml) as Record<string, any>
    const lines: string[] = [`# ${String(sheet['@_name'] ?? '')}`]
    for (const row of asArray(worksheet.worksheet?.sheetData?.row) as Array<
      Record<string, unknown>
    >) {
      const cells: string[] = []
      for (const cell of asArray(row.c as Cell | Cell[])) {
        const text = cellText(cell, shared)
        const col = cell['@_r'] ? columnIndex(cell['@_r']) : cells.length
        while (cells.length < col) cells.push('')
        cells[col] = text
      }
      lines.push(cells.join(' | '))
    }
    sections.push(lines.join('\n'))
  }
  return sections.join('\n\n')
}
