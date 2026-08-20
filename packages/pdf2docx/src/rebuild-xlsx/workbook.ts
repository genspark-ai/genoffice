/**
 * Minimal xlsx packaging (P26): a fresh workbook written from scratch —
 * [Content_Types].xml, _rels, workbook.xml, one worksheet per page, and a
 * deduplicated styles.xml (numFmt/font/fill/border/alignment pools). Strings
 * ship as inline strings (no sharedStrings part — simpler, and every string
 * is unique table text anyway). Cell XML shapes follow apps/sheets'
 * xlsx-gateway serializer so both readers agree on the dialect.
 */
import JSZip from 'jszip'
import { BUILTIN_NUM_FMTS } from './numbers'

// ── style model ──

export interface FontSpec {
  bold?: boolean
  italic?: boolean
  /** hex RRGGBB without '#' */
  color?: string
  /** font size in points (default 11) */
  sizePt?: number
  name?: string
}

export interface BorderEdges {
  top?: string
  right?: string
  bottom?: string
  left?: string
}

export interface AlignmentSpec {
  vertical?: 'top' | 'center' | 'bottom'
  horizontal?: 'left' | 'center' | 'right'
  wrapText?: boolean
}

export interface CellStyleSpec {
  /** numFmt code ('General' or e.g. '0.0%'); builtin codes map to builtin ids */
  numFmt?: string
  font?: FontSpec
  /** solid fill, hex RRGGBB without '#' */
  fill?: string
  /** thin border edges, each a hex RRGGBB color */
  border?: BorderEdges
  align?: AlignmentSpec
}

// ── sheet model ──

export type CellValue = { kind: 'text'; text: string } | { kind: 'number'; value: number }

export interface SheetCell {
  /** 0-based */
  row: number
  /** 0-based */
  col: number
  value?: CellValue
  /** index from StylePool.cellXf; 0 = default */
  styleId: number
}

export interface SheetSpec {
  name: string
  cells: SheetCell[]
  /** per-column width in Excel character units; sparse (undefined = default) */
  colWidths?: Array<number | undefined>
  /** measured row heights in points, by 0-based row index */
  rowHeightsPt?: Map<number, number>
  /** merged ranges in A1 notation ("A1:B2") */
  merges?: string[]
}

// ── style pool ──

const FIRST_CUSTOM_NUM_FMT = 164

/**
 * Deduplicating style pool: identical specs share one cellXfs entry. Slot 0
 * of each pool is the default (empty font / no fill / no border / xf 0), and
 * fill slot 1 stays the gray125 pattern both Excel and LO expect to exist.
 */
export class StylePool {
  private numFmts = new Map<string, number>()
  private fonts = new Map<string, number>(['{}'].map((k, i) => [k, i]))
  private fills = new Map<string, number>([
    ['none', 0],
    ['gray125', 1],
  ])
  private borders = new Map<string, number>([['{}', 0]])
  private xfs = new Map<string, number>()
  private xfEntries: string[] = []

  constructor() {
    // xf 0 must exist and stay all-defaults (rows/columns reference it)
    this.cellXf({})
  }

  private numFmtId(code: string): number {
    const builtin = BUILTIN_NUM_FMTS[code]
    if (builtin !== undefined) return builtin
    let id = this.numFmts.get(code)
    if (id === undefined) {
      id = FIRST_CUSTOM_NUM_FMT + this.numFmts.size
      this.numFmts.set(code, id)
    }
    return id
  }

  private fontId(font: FontSpec | undefined): number {
    const key = JSON.stringify(font ?? {})
    let id = this.fonts.get(key)
    if (id === undefined) {
      id = this.fonts.size
      this.fonts.set(key, id)
    }
    return id
  }

  private fillId(fill: string | undefined): number {
    const key = fill ?? 'none'
    let id = this.fills.get(key)
    if (id === undefined) {
      id = this.fills.size
      this.fills.set(key, id)
    }
    return id
  }

  private borderId(border: BorderEdges | undefined): number {
    const key = JSON.stringify(border ?? {})
    let id = this.borders.get(key)
    if (id === undefined) {
      id = this.borders.size
      this.borders.set(key, id)
    }
    return id
  }

  /** resolve a style spec to its (deduplicated) cellXfs index */
  cellXf(spec: CellStyleSpec): number {
    const numFmtId = this.numFmtId(spec.numFmt ?? 'General')
    const fontId = this.fontId(spec.font)
    const fillId = this.fillId(spec.fill)
    const borderId = this.borderId(spec.border)
    const align = spec.align ?? {}
    const alignParts: string[] = []
    if (align.horizontal) alignParts.push(`horizontal="${align.horizontal}"`)
    if (align.vertical) {
      alignParts.push(`vertical="${align.vertical === 'center' ? 'center' : align.vertical}"`)
    }
    if (align.wrapText) alignParts.push('wrapText="1"')
    const alignXml = alignParts.length > 0 ? `<alignment ${alignParts.join(' ')}/>` : ''

    const applies =
      (numFmtId !== 0 ? ' applyNumberFormat="1"' : '') +
      (fontId !== 0 ? ' applyFont="1"' : '') +
      (fillId > 1 ? ' applyFill="1"' : '') +
      (borderId !== 0 ? ' applyBorder="1"' : '') +
      (alignXml !== '' ? ' applyAlignment="1"' : '')
    const attrs = `numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${applies}`
    const entry = alignXml === '' ? `<xf ${attrs}/>` : `<xf ${attrs}>${alignXml}</xf>`
    let id = this.xfs.get(entry)
    if (id === undefined) {
      id = this.xfEntries.length
      this.xfs.set(entry, id)
      this.xfEntries.push(entry)
    }
    return id
  }

  stylesXml(): string {
    const numFmtsXml =
      this.numFmts.size === 0
        ? ''
        : `<numFmts count="${this.numFmts.size}">` +
          [...this.numFmts.entries()]
            .map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`)
            .join('') +
          '</numFmts>'

    const fontsXml = [...this.fonts.keys()]
      .map((key) => {
        const font = JSON.parse(key) as FontSpec
        const parts: string[] = []
        if (font.bold) parts.push('<b/>')
        if (font.italic) parts.push('<i/>')
        parts.push(`<sz val="${font.sizePt ?? 11}"/>`)
        if (font.color && /^[0-9a-fA-F]{6}$/.test(font.color)) {
          parts.push(`<color rgb="FF${font.color.toUpperCase()}"/>`)
        }
        parts.push(`<name val="${escapeXml(font.name ?? 'Calibri')}"/>`)
        return `<font>${parts.join('')}</font>`
      })
      .join('')

    const fillsXml = [...this.fills.keys()]
      .map((key) => {
        if (key === 'none') return '<fill><patternFill patternType="none"/></fill>'
        if (key === 'gray125') return '<fill><patternFill patternType="gray125"/></fill>'
        return (
          '<fill><patternFill patternType="solid">' +
          `<fgColor rgb="FF${key.toUpperCase()}"/><bgColor rgb="FF${key.toUpperCase()}"/>` +
          '</patternFill></fill>'
        )
      })
      .join('')

    const bordersXml = [...this.borders.keys()]
      .map((key) => {
        const edges = JSON.parse(key) as BorderEdges
        const edge = (side: 'left' | 'right' | 'top' | 'bottom'): string => {
          const color = edges[side]
          if (!color) return `<${side}/>`
          return `<${side} style="thin"><color rgb="FF${color.toUpperCase()}"/></${side}>`
        }
        return `<border>${edge('left')}${edge('right')}${edge('top')}${edge('bottom')}<diagonal/></border>`
      })
      .join('')

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmtsXml +
      `<fonts count="${this.fonts.size}">${fontsXml}</fonts>` +
      `<fills count="${this.fills.size}">${fillsXml}</fills>` +
      `<borders count="${this.borders.size}">${bordersXml}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfEntries.length}">${this.xfEntries.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>'
    )
  }
}

// ── XML helpers ──

export function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/** control chars xlsx forbids in text (tab/newline/CR survive as literals) */
function sanitizeText(value: string): string {
  // eslint-disable-next-line no-control-regex
  return value.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
}

/** 0-based column index → A1 letters */
export function columnLabel(column: number): string {
  let label = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    label = String.fromCharCode(65 + (remaining % 26)) + label
    remaining = Math.floor(remaining / 26)
  }
  return label
}

export function cellRef(row: number, col: number): string {
  return `${columnLabel(col)}${row + 1}`
}

// ── worksheet serialization ──

function cellXml(cell: SheetCell): string {
  const ref = cellRef(cell.row, cell.col)
  const style = cell.styleId !== 0 ? ` s="${cell.styleId}"` : ''
  if (!cell.value) return `<c r="${ref}"${style}/>`
  if (cell.value.kind === 'number') return `<c r="${ref}"${style}><v>${cell.value.value}</v></c>`
  return (
    `<c r="${ref}"${style} t="inlineStr">` +
    `<is><t xml:space="preserve">${escapeXml(sanitizeText(cell.value.text))}</t></is></c>`
  )
}

export function worksheetXml(sheet: SheetSpec): string {
  const byRow = new Map<number, SheetCell[]>()
  let maxRow = 0
  let maxCol = 0
  for (const cell of sheet.cells) {
    maxRow = Math.max(maxRow, cell.row)
    maxCol = Math.max(maxCol, cell.col)
    const row = byRow.get(cell.row)
    if (row) row.push(cell)
    else byRow.set(cell.row, [cell])
  }
  for (const [, width] of (sheet.colWidths ?? []).entries()) {
    if (width !== undefined) maxCol = Math.max(maxCol, 0)
  }

  const rowsXml = [...byRow.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([rowIdx, cells]) => {
      const heightPt = sheet.rowHeightsPt?.get(rowIdx)
      const height =
        heightPt !== undefined ? ` ht="${Number(heightPt.toFixed(2))}" customHeight="1"` : ''
      const cellsXml = cells
        .sort((a, b) => a.col - b.col)
        .map(cellXml)
        .join('')
      return `<row r="${rowIdx + 1}"${height}>${cellsXml}</row>`
    })
    .join('')

  const colsEntries = (sheet.colWidths ?? [])
    .map((width, i) =>
      width === undefined
        ? ''
        : `<col min="${i + 1}" max="${i + 1}" width="${Number(width.toFixed(2))}" customWidth="1"/>`,
    )
    .join('')
  const colsXml = colsEntries === '' ? '' : `<cols>${colsEntries}</cols>`

  const merges = sheet.merges ?? []
  const mergesXml =
    merges.length === 0
      ? ''
      : `<mergeCells count="${merges.length}">` +
        merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('') +
        '</mergeCells>'

  const dimension = `A1:${cellRef(Math.max(maxRow, 0), Math.max(maxCol, 0))}`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<dimension ref="${dimension}"/>` +
    colsXml +
    `<sheetData>${rowsXml}</sheetData>` +
    mergesXml +
    '</worksheet>'
  )
}

// ── package assembly ──

export async function buildXlsxPackage(
  sheets: SheetSpec[],
  styles: StylePool,
): Promise<Uint8Array> {
  const zip = new JSZip()
  const n = sheets.length

  const sheetOverrides = sheets
    .map(
      (_, i) =>
        `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ` +
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    )
    .join('')
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheetOverrides +
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

  const sheetEntries = sheets
    .map(
      (sheet, i) =>
        `<sheet name="${escapeXml(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`,
    )
    .join('')
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets>${sheetEntries}</sheets></workbook>`,
  )

  const relEntries = sheets
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" ` +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ' +
        `Target="worksheets/sheet${i + 1}.xml"/>`,
    )
    .join('')
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      relEntries +
      `<Relationship Id="rId${n + 1}" ` +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" ' +
      'Target="styles.xml"/>' +
      '</Relationships>',
  )

  zip.file('xl/styles.xml', styles.stylesXml())
  sheets.forEach((sheet, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, worksheetXml(sheet)))

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
