/// Surgical hyperlink editing for a worksheet part and its .rels sibling.
/// External targets live as TargetMode="External" relationships referenced by
/// r:id; internal anchors use the `location` attribute and need no rel.

export { ensureRelationshipNamespace } from './xlsx-namespace'

export class HyperlinkEditError extends Error {}

export interface HyperlinkEdit {
  readonly row: number
  readonly column: number
  /// '#Sheet!A1' = internal anchor (leading '#' stripped into `location`);
  /// anything else = external URL; null removes the link.
  readonly target: string | null
}

export interface HyperlinkPatch {
  readonly worksheetXml: string
  /// null when the sheet had no rels file and still needs none.
  readonly relsXml: string | null
  readonly relsChanged: boolean
}

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>'

export function applyHyperlinkEdits(
  worksheetXml: string,
  relsXml: string | null,
  edits: readonly HyperlinkEdit[],
): HyperlinkPatch {
  let xml = worksheetXml
  let rels = relsXml
  let relsChanged = false

  const relIdsInUse = (): Set<string> => {
    const used = new Set<string>()
    for (const element of xml.matchAll(/<hyperlink\b[^>]*?\br:id="([^"]+)"/g)) {
      if (element[1]) used.add(element[1])
    }
    return used
  }
  const dropUnusedRel = (relId: string): void => {
    if (rels === null || relIdsInUse().has(relId)) return
    const next = rels.replace(new RegExp(`<Relationship\\b[^>]*\\bId="${relId}"[^>]*/>`), '')
    if (next !== rels) {
      rels = next
      relsChanged = true
    }
  }
  const allocateRel = (target: string): string => {
    if (rels === null) {
      rels = EMPTY_RELS
      relsChanged = true
    }
    let maximum = 0
    for (const id of rels.matchAll(/\bId="rId([0-9]+)"/g)) {
      maximum = Math.max(maximum, Number(id[1]))
    }
    const relId = `rId${maximum + 1}`
    const element =
      `<Relationship Id="${relId}" Type="${HYPERLINK_REL_TYPE}" ` +
      `Target="${escapeXmlAttribute(target)}" TargetMode="External"/>`
    rels = rels.replace('</Relationships>', () => `${element}</Relationships>`)
    relsChanged = true
    return relId
  }

  for (const edit of edits) {
    const ref = toA1(edit.row, edit.column)
    const existingPattern = new RegExp(`<hyperlink\\b[^>]*?\\bref="${ref}"[^>]*/>`)
    const existing = existingPattern.exec(xml)?.[0]
    const existingRelId =
      existing === undefined ? undefined : /\br:id="([^"]+)"/.exec(existing)?.[1]
    if (existing !== undefined) {
      xml = xml.replace(existingPattern, '')
    }
    if (existingRelId !== undefined) dropUnusedRel(existingRelId)
    if (edit.target === null) continue

    const element = edit.target.startsWith('#')
      ? `<hyperlink ref="${ref}" location="${escapeXmlAttribute(edit.target.slice(1))}"/>`
      : `<hyperlink ref="${ref}" r:id="${allocateRel(edit.target)}"/>`
    xml = insertHyperlinkElement(xml, element)
  }

  // An emptied section must go — Excel repairs `<hyperlinks/>` with no children.
  xml = xml.replace(/<hyperlinks>\s*<\/hyperlinks>/, '')
  return { worksheetXml: xml, relsXml: rels, relsChanged }
}

function insertHyperlinkElement(xml: string, element: string): string {
  if (xml.includes('</hyperlinks>')) {
    return xml.replace('</hyperlinks>', () => `${element}</hyperlinks>`)
  }
  const section = `<hyperlinks>${element}</hyperlinks>`
  // Schema order: hyperlinks follows dataValidations / conditionalFormatting /
  // mergeCells / sheetData and precedes print-related elements and drawings.
  const anchor =
    /<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/.exec(
      xml,
    )
  if (anchor) {
    return xml.slice(0, anchor.index) + section + xml.slice(anchor.index)
  }
  const end = xml.lastIndexOf('</worksheet>')
  if (end === -1) throw new HyperlinkEditError('Worksheet has no closing element.')
  return xml.slice(0, end) + section + xml.slice(end)
}

function toA1(row: number, column: number): string {
  let letters = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return `${letters}${row + 1}`
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}
