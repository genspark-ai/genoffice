/// x14 sparkline writer: appends sparkline groups to a worksheet's extLst
/// (creating extLst / the x14 ext as needed). Existing groups stay verbatim.

export class SparklineAddError extends Error {}

export interface SparklineCellAdd {
  readonly cell: string
  /// Sheet-qualified source range, e.g. `Sheet1!A2:F2`.
  readonly sourceRef: string
}

export interface SparklineGroupAdd {
  readonly type: 'line' | 'column' | 'stacked'
  readonly color?: string | undefined
  readonly cells: readonly SparklineCellAdd[]
}

export interface SheetSparklineAddition extends SparklineGroupAdd {
  readonly sheetName: string
}

const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main'
const DEFAULT_SERIES_ARGB = 'FF376092'
const NEGATIVE_ARGB = 'FFD00000'

function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function toArgb(color: string | undefined): string {
  if (color === undefined) return DEFAULT_SERIES_ARGB
  if (!/^#[0-9A-Fa-f]{6}$/.test(color)) {
    throw new SparklineAddError(`"${color}" is not a #RRGGBB color.`)
  }
  return `FF${color.slice(1).toUpperCase()}`
}

function buildGroupXml(group: SparklineGroupAdd): string {
  const typeAttribute = group.type === 'line' ? '' : ` type="${group.type}"`
  const sparklines = group.cells
    .map(
      (cell) =>
        `<x14:sparkline><xm:f>${escapeXml(cell.sourceRef)}</xm:f>` +
        `<xm:sqref>${cell.cell}</xm:sqref></x14:sparkline>`,
    )
    .join('')
  return (
    `<x14:sparklineGroup displayEmptyCellsAs="gap"${typeAttribute}>` +
    `<x14:colorSeries rgb="${toArgb(group.color)}"/>` +
    `<x14:colorNegative rgb="${NEGATIVE_ARGB}"/>` +
    `<x14:sparklines>${sparklines}</x14:sparklines></x14:sparklineGroup>`
  )
}

function wrapGroups(groupsXml: string): string {
  return `<x14:sparklineGroups xmlns:xm="${XM_NS}">${groupsXml}</x14:sparklineGroups>`
}

function insertGroups(worksheetXml: string, groupsXml: string): string {
  const extOpen = /<ext\b[^>]*\buri="\{05C60535-1F16-4FD2-B633-F4F36F0B64E0\}"[^>]*>/i.exec(
    worksheetXml,
  )
  if (extOpen && !extOpen[0].endsWith('/>')) {
    const closeAt = worksheetXml.indexOf('</ext>', extOpen.index)
    if (closeAt === -1) throw new SparklineAddError('The sparkline ext element is malformed.')
    const block = worksheetXml.slice(extOpen.index, closeAt)
    const groupsClose = block.lastIndexOf('</x14:sparklineGroups>')
    const patched =
      groupsClose === -1
        ? block + wrapGroups(groupsXml)
        : block.slice(0, groupsClose) + groupsXml + block.slice(groupsClose)
    return worksheetXml.slice(0, extOpen.index) + patched + worksheetXml.slice(closeAt)
  }
  const ext =
    `<ext uri="${SPARKLINE_EXT_URI}" xmlns:x14="${X14_NS}">` + `${wrapGroups(groupsXml)}</ext>`
  // Only the worksheet-level extLst (past sheetData) qualifies — cells may
  // carry their own extLst.
  const extLstClose = worksheetXml.lastIndexOf('</extLst>')
  if (extLstClose !== -1 && extLstClose > worksheetXml.lastIndexOf('</sheetData>')) {
    return worksheetXml.slice(0, extLstClose) + ext + worksheetXml.slice(extLstClose)
  }
  const worksheetClose = worksheetXml.lastIndexOf('</worksheet>')
  if (worksheetClose === -1) throw new SparklineAddError('Worksheet has no closing element.')
  return (
    worksheetXml.slice(0, worksheetClose) +
    `<extLst>${ext}</extLst>` +
    worksheetXml.slice(worksheetClose)
  )
}

/// Appends one x14:sparklineGroup per addition. Fails closed when a host
/// cell already carries a sparkline (in the file or within this save).
export function applySparklineAdditions(
  worksheetXml: string,
  additions: readonly SparklineGroupAdd[],
): string {
  if (additions.length === 0) return worksheetXml
  const seen = new Set<string>()
  for (const group of additions) {
    for (const { cell } of group.cells) {
      if (seen.has(cell) || worksheetXml.includes(`<xm:sqref>${cell}</xm:sqref>`)) {
        throw new SparklineAddError(`${cell} already has a sparkline.`)
      }
      seen.add(cell)
    }
  }
  return insertGroups(worksheetXml, additions.map(buildGroupXml).join(''))
}
