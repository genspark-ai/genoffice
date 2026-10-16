/// Persists "picture in cell" images created in the editor as Excel's native
/// rich-value parts — xl/metadata.xml (an XLRICHVALUE metadata type plus one
/// futureMetadata/valueMetadata record per image), xl/richData/
/// rdrichvaluestructure.xml, rdrichvalue.xml, richValueRel.xml and its
/// relationships to the media, the workbook relationships, and the
/// [Content_Types].xml overrides — then marks each host cell with `vm=` and
/// the cached #VALUE! Excel writes for "Place in Cell". Reopening resolves
/// the same parts through the sidecar's richdata reader, and recent WPS
/// builds display Excel pictures in cells.

import {
  allocatePartPath,
  appendRelationship,
  registerContentTypeDefault,
  registerContentTypeOverride,
  relativeTarget,
  relsPathFor,
  type MutablePackage,
} from './xlsx-drawing-add'

export class CellImageSaveError extends Error {}

export interface CellImageAddition {
  readonly worksheetPath: string
  /// 0-based target cell; the save writes the picture's rich value there.
  readonly row: number
  readonly column: number
  readonly mediaType: 'image/png' | 'image/jpeg' | 'image/gif'
  readonly base64: string
}

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const RICH_DATA_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata'
const RICH_VALUE_REL_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel'
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const METADATA_PATH = 'xl/metadata.xml'
const METADATA_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'
const METADATA_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata'
const STRUCTURE_PATH = 'xl/richData/rdrichvaluestructure.xml'
const STRUCTURE_CONTENT_TYPE = 'application/vnd.ms-excel.rdrichvaluestructure+xml'
const STRUCTURE_REL_TYPE =
  'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure'
const RICH_VALUE_PATH = 'xl/richData/rdrichvalue.xml'
const RICH_VALUE_CONTENT_TYPE = 'application/vnd.ms-excel.rdrichvalue+xml'
const RICH_VALUE_REL_TYPE = 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue'
const RICH_VALUE_REL_PATH = 'xl/richData/richValueRel.xml'
const RICH_VALUE_REL_CONTENT_TYPE = 'application/vnd.ms-excel.richvaluerel+xml'
const RICH_VALUE_REL_REL_TYPE =
  'http://schemas.microsoft.com/office/2022/10/relationships/richValueRel'
const IMAGE_REL_TYPE = `${OFFICE_REL_NS}/image`
/// Excel's own capability flags for the XLRICHVALUE metadata type.
const RICH_VALUE_METADATA_TYPE =
  '<metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" ' +
  'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" ' +
  'clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
/// The extension Excel wraps each `xlrd:rvb` record in.
const RICH_VALUE_BK_EXT_URI = '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}'
/// The rich value's `CalcOrigin` for a picture inserted from a file.
const CALC_ORIGIN_LOCAL_IMAGE = '5'
const IMAGE_EXTENSIONS: Record<CellImageAddition['mediaType'], string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
}

/// Writes each addition's media, rich-value bookkeeping, and `vm=`-marked
/// host cell. Parts that already carry rich data (Excel files with pictures
/// in cells, or a dynamic-array metadata part) are extended in place so the
/// existing records keep their indexes.
export async function applyCellImageAdditions(
  pkg: MutablePackage,
  additions: readonly CellImageAddition[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (additions.length === 0) return

  // Media parts first: their paths go into richValueRel's relationships.
  const mediaPaths: string[] = []
  for (const addition of additions) {
    const extension = IMAGE_EXTENSIONS[addition.mediaType]
    const mediaPath = await allocatePartPath(pkg, 'xl/media/image', `.${extension}`)
    pkg.addBinary(mediaPath, Uint8Array.from(Buffer.from(addition.base64, 'base64')))
    touchedEntries.add(mediaPath)
    mediaPaths.push(mediaPath)
    await registerContentTypeDefault(pkg, extension, addition.mediaType, touchedEntries)
  }

  // Index bookkeeping is read before anything appends: rdrichvalue entries
  // and richValueRel entries pair up by position, and the metadata records
  // point back at those positions.
  const richValueBase = await countPartElements(pkg, RICH_VALUE_PATH, /<rv\b/g)
  const relBase = await countPartElements(pkg, RICH_VALUE_REL_PATH, /<rel\b/g)
  const structureIndex = await ensureLocalImageStructure(pkg, touchedEntries)
  const richValueIndexes = additions.map((_, at) => richValueBase + at)

  await appendRichValues(
    pkg,
    richValueIndexes.map((_, at) => ({ structureIndex, relIndex: relBase + at })),
    touchedEntries,
  )
  await appendRichValueRels(pkg, mediaPaths, touchedEntries)
  const vmIndexes = await appendRichValueMetadata(pkg, richValueIndexes, touchedEntries)
  await ensureWorkbookRelationships(pkg, touchedEntries)
  await registerContentTypeOverride(pkg, METADATA_PATH, METADATA_CONTENT_TYPE, touchedEntries)
  await registerContentTypeOverride(pkg, STRUCTURE_PATH, STRUCTURE_CONTENT_TYPE, touchedEntries)
  await registerContentTypeOverride(pkg, RICH_VALUE_PATH, RICH_VALUE_CONTENT_TYPE, touchedEntries)
  await registerContentTypeOverride(
    pkg,
    RICH_VALUE_REL_PATH,
    RICH_VALUE_REL_CONTENT_TYPE,
    touchedEntries,
  )

  // Cell patches run last and against the package's final worksheet XML.
  const byWorksheet = new Map<string, { row: number; column: number; vm: number }[]>()
  additions.forEach((addition, at) => {
    const images = byWorksheet.get(addition.worksheetPath) ?? []
    images.push({ row: addition.row, column: addition.column, vm: vmIndexes[at] ?? at + 1 })
    byWorksheet.set(addition.worksheetPath, images)
  })
  for (const [worksheetPath, images] of byWorksheet) {
    let xml = await pkg.readText(worksheetPath)
    for (const image of images) {
      xml = patchCellAsRichImage(xml, toA1Address(image.row, image.column), image.vm)
    }
    pkg.write(worksheetPath, xml)
    touchedEntries.add(worksheetPath)
  }
}

async function countPartElements(
  pkg: MutablePackage,
  path: string,
  pattern: RegExp,
): Promise<number> {
  if (!(await pkg.has(path))) return 0
  return [...(await pkg.readText(path)).matchAll(pattern)].length
}

/// The index of the `_localImage` rich-value structure (Excel's shape for a
/// picture in a cell), appending the structure when the part lacks one.
async function ensureLocalImageStructure(
  pkg: MutablePackage,
  touchedEntries: Set<string>,
): Promise<number> {
  const structure =
    '<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s>'
  if (!(await pkg.has(STRUCTURE_PATH))) {
    pkg.add(
      STRUCTURE_PATH,
      `${XML_HEADER}<rvStructures xmlns="${RICH_DATA_NS}" count="1">${structure}</rvStructures>`,
    )
    touchedEntries.add(STRUCTURE_PATH)
    return 0
  }
  const xml = await pkg.readText(STRUCTURE_PATH)
  const structures = [...xml.matchAll(/<s\b[^>]*>/g)]
  const existing = structures.findIndex((match) => /\bt="_localImage"/.test(match[0]))
  if (existing >= 0) return existing
  const open = /<rvStructures\b[^>]*>/.exec(xml)?.[0]
  if (!open) throw new CellImageSaveError('rdrichvaluestructure.xml is malformed.')
  const patched = bumpCount(
    xml.replace('</rvStructures>', () => `${structure}</rvStructures>`),
    open,
    1,
  )
  pkg.write(STRUCTURE_PATH, patched)
  touchedEntries.add(STRUCTURE_PATH)
  return structures.length
}

/// One `<rv>` per picture: the structure index, the richValueRel position,
/// and CalcOrigin.
async function appendRichValues(
  pkg: MutablePackage,
  entries: readonly { readonly structureIndex: number; readonly relIndex: number }[],
  touchedEntries: Set<string>,
): Promise<void> {
  const elements = entries
    .map(
      ({ structureIndex, relIndex }) =>
        `<rv s="${structureIndex}"><v>${relIndex}</v><v>${CALC_ORIGIN_LOCAL_IMAGE}</v></rv>`,
    )
    .join('')
  if (!(await pkg.has(RICH_VALUE_PATH))) {
    pkg.add(
      RICH_VALUE_PATH,
      `${XML_HEADER}<rvData xmlns="${RICH_DATA_NS}" count="${entries.length}">${elements}</rvData>`,
    )
    touchedEntries.add(RICH_VALUE_PATH)
    return
  }
  const xml = await pkg.readText(RICH_VALUE_PATH)
  const open = /<rvData\b[^>]*>/.exec(xml)?.[0]
  if (!open) throw new CellImageSaveError('rdrichvalue.xml is malformed.')
  const patched = bumpCount(
    xml.replace('</rvData>', () => `${elements}</rvData>`),
    open,
    entries.length,
  )
  pkg.write(RICH_VALUE_PATH, patched)
  touchedEntries.add(RICH_VALUE_PATH)
}

/// richValueRel.xml lists the media in order; the relationship file behind it
/// (`xl/richData/_rels/richValueRel.xml.rels`) resolves each entry's id.
async function appendRichValueRels(
  pkg: MutablePackage,
  mediaPaths: readonly string[],
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = relsPathFor(RICH_VALUE_REL_PATH)
  const elements: string[] = []
  for (const mediaPath of mediaPaths) {
    const id = await appendRelationship(
      pkg,
      relsPath,
      IMAGE_REL_TYPE,
      relativeTarget(RICH_VALUE_REL_PATH, mediaPath),
    )
    elements.push(`<rel r:id="${id}"/>`)
  }
  touchedEntries.add(relsPath)
  if (!(await pkg.has(RICH_VALUE_REL_PATH))) {
    pkg.add(
      RICH_VALUE_REL_PATH,
      `${XML_HEADER}<richValueRels xmlns="${RICH_VALUE_REL_NS}" xmlns:r="${OFFICE_REL_NS}">${elements.join('')}</richValueRels>`,
    )
  } else {
    const xml = await pkg.readText(RICH_VALUE_REL_PATH)
    pkg.write(
      RICH_VALUE_REL_PATH,
      xml.replace('</richValueRels>', () => `${elements.join('')}</richValueRels>`),
    )
  }
  touchedEntries.add(RICH_VALUE_REL_PATH)
}

/// Appends one futureMetadata/valueMetadata record per picture and returns
/// each host cell's 1-based `vm` index. A part without rich values yet (a
/// dynamic-array metadata.xml, or none at all) is created or extended.
async function appendRichValueMetadata(
  pkg: MutablePackage,
  richValueIndexes: readonly number[],
  touchedEntries: Set<string>,
): Promise<number[]> {
  if (!(await pkg.has(METADATA_PATH))) {
    pkg.add(METADATA_PATH, buildMetadataXml(richValueIndexes))
    touchedEntries.add(METADATA_PATH)
    return richValueIndexes.map((_, at) => at + 1)
  }
  const patched = appendMetadataRecords(await pkg.readText(METADATA_PATH), richValueIndexes)
  pkg.write(METADATA_PATH, patched.xml)
  touchedEntries.add(METADATA_PATH)
  return patched.vmIndexes
}

function buildMetadataXml(richValueIndexes: readonly number[]): string {
  const future = richValueIndexes.map(richValueBk).join('')
  const values = richValueIndexes.map((_, at) => `<bk><rc t="1" v="${at}"/></bk>`).join('')
  return (
    XML_HEADER +
    `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="${RICH_DATA_NS}">` +
    `<metadataTypes count="1">${RICH_VALUE_METADATA_TYPE}</metadataTypes>` +
    `<futureMetadata name="XLRICHVALUE" count="${richValueIndexes.length}">${future}</futureMetadata>` +
    `<valueMetadata count="${richValueIndexes.length}">${values}</valueMetadata>` +
    '</metadata>'
  )
}

/// Extends an existing metadata part (XLDAPR-only, or rich values written by
/// Excel) with this save's records, keeping every earlier index stable.
function appendMetadataRecords(
  xml: string,
  richValueIndexes: readonly number[],
): { readonly xml: string; readonly vmIndexes: number[] } {
  let patched = ensureRichValueNamespace(xml)
  const typesBlock = /<metadataTypes\b[^>]*>([\s\S]*?)<\/metadataTypes>/.exec(patched)
  const typesOpen = typesBlock ? /<metadataTypes\b[^>]*>/.exec(typesBlock[0])?.[0] : undefined
  if (!typesBlock || !typesOpen) throw new CellImageSaveError('xl/metadata.xml is malformed.')
  const typeNames = [...typesBlock[1]!.matchAll(/<metadataType\b[^>]*\bname="([^"]*)"/g)].map(
    (match) => match[1],
  )
  let typeIndex = typeNames.indexOf('XLRICHVALUE') + 1
  if (typeIndex === 0) {
    const withType = typesBlock[0].replace(
      '</metadataTypes>',
      () => `${RICH_VALUE_METADATA_TYPE}</metadataTypes>`,
    )
    patched = patched.replace(typesBlock[0], () => bumpCount(withType, typesOpen, 1))
    typeIndex = typeNames.length + 1
  }

  const futureBlock =
    /<futureMetadata\b[^>]*\bname="XLRICHVALUE"[^>]*>([\s\S]*?)<\/futureMetadata>/.exec(patched)
  let futureBase = 0
  if (futureBlock) {
    futureBase = [...futureBlock[1]!.matchAll(/<bk\b/g)].length
    const open = /<futureMetadata\b[^>]*>/.exec(futureBlock[0])?.[0]
    if (!open) throw new CellImageSaveError('xl/metadata.xml is malformed.')
    const records = richValueIndexes.map(richValueBk).join('')
    const withRecords = futureBlock[0].replace(
      '</futureMetadata>',
      () => `${records}</futureMetadata>`,
    )
    patched = patched.replace(futureBlock[0], () =>
      bumpCount(withRecords, open, richValueIndexes.length),
    )
  } else {
    const records = richValueIndexes.map(richValueBk).join('')
    const element = `<futureMetadata name="XLRICHVALUE" count="${richValueIndexes.length}">${records}</futureMetadata>`
    patched = insertBeforeFirst(
      patched,
      ['<cellMetadata', '<valueMetadata', '</metadata>'],
      element,
    )
  }

  const valuesBlock = /<valueMetadata\b[^>]*>([\s\S]*?)<\/valueMetadata>/.exec(patched)
  let vmBase = 0
  if (valuesBlock) {
    vmBase = [...valuesBlock[1]!.matchAll(/<bk\b/g)].length
    const open = /<valueMetadata\b[^>]*>/.exec(valuesBlock[0])?.[0]
    if (!open) throw new CellImageSaveError('xl/metadata.xml is malformed.')
    const records = richValueIndexes
      .map((_, at) => `<bk><rc t="${typeIndex}" v="${futureBase + at}"/></bk>`)
      .join('')
    const withRecords = valuesBlock[0].replace(
      '</valueMetadata>',
      () => `${records}</valueMetadata>`,
    )
    patched = patched.replace(valuesBlock[0], () =>
      bumpCount(withRecords, open, richValueIndexes.length),
    )
  } else {
    const records = richValueIndexes
      .map((_, at) => `<bk><rc t="${typeIndex}" v="${futureBase + at}"/></bk>`)
      .join('')
    const element = `<valueMetadata count="${richValueIndexes.length}">${records}</valueMetadata>`
    patched = patched.replace('</metadata>', () => `${element}</metadata>`)
  }

  return { xml: patched, vmIndexes: richValueIndexes.map((_, at) => vmBase + at + 1) }
}

function richValueBk(richValueIndex: number): string {
  return (
    `<bk><extLst><ext uri="${RICH_VALUE_BK_EXT_URI}">` +
    `<xlrd:rvb i="${richValueIndex}"/></ext></extLst></bk>`
  )
}

/// The `xlrd` prefix the rvb records need; Excel's own parts declare it on
/// the root, a dynamic-array metadata part written here does not.
function ensureRichValueNamespace(xml: string): string {
  const open = /<metadata\b[^>]*>/.exec(xml)?.[0]
  if (!open || /xmlns:xlrd=/.test(open)) return xml
  return xml.replace(open, () => open.replace(/>$/, ` xmlns:xlrd="${RICH_DATA_NS}">`))
}

/// Adds `delta` to the `count=` attribute of `openTag` inside `xml`.
function bumpCount(xml: string, openTag: string, delta: number): string {
  const next = /\bcount="(\d+)"/.test(openTag)
    ? openTag.replace(/\bcount="(\d+)"/, (_, value: string) => `count="${Number(value) + delta}"`)
    : openTag.replace(/>$/, ` count="${delta}">`)
  return xml.replace(openTag, () => next)
}

/// Inserts `element` before the first of `markers` that appears in `xml`.
function insertBeforeFirst(xml: string, markers: readonly string[], element: string): string {
  let at = -1
  for (const marker of markers) {
    const index = xml.indexOf(marker)
    if (index >= 0 && (at === -1 || index < at)) at = index
  }
  if (at === -1) throw new CellImageSaveError('xl/metadata.xml is malformed.')
  return xml.slice(0, at) + element + xml.slice(at)
}

/// The workbook part must point at the metadata and rich-data parts; Excel
/// resolves them by relationship type, so existing ones are kept as-is.
async function ensureWorkbookRelationships(
  pkg: MutablePackage,
  touchedEntries: Set<string>,
): Promise<void> {
  const path = 'xl/_rels/workbook.xml.rels'
  const xml = await pkg.readText(path)
  const wanted: readonly (readonly [string, string])[] = [
    [METADATA_REL_TYPE, 'metadata.xml'],
    [RICH_VALUE_REL_REL_TYPE, 'richData/richValueRel.xml'],
    [RICH_VALUE_REL_TYPE, 'richData/rdrichvalue.xml'],
    [STRUCTURE_REL_TYPE, 'richData/rdrichvaluestructure.xml'],
  ]
  for (const [type, target] of wanted) {
    if (xml.includes(`Type="${type}"`)) continue
    await appendRelationship(pkg, path, type, target)
    touchedEntries.add(path)
  }
}

/// Rewrites one cell as an in-cell picture: the style survives, the value
/// (or formula) is replaced by the cached #VALUE! Excel writes, and the
/// `vm` index points at this save's valueMetadata record.
function patchCellAsRichImage(worksheetXml: string, address: string, vm: number): string {
  const cellPattern = new RegExp(`<c\\b([^>]*?)\\br="${address}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`)
  const existing = cellPattern.exec(worksheetXml)
  const style = existing
    ? readXmlAttribute(`${existing[1] ?? ''} ${existing[2] ?? ''}`, 's')
    : undefined
  const styleAttribute = style === undefined ? '' : ` s="${style}"`
  const replacement = `<c r="${address}"${styleAttribute} t="e" vm="${vm}"><v>#VALUE!</v></c>`
  if (existing) return worksheetXml.replace(cellPattern, () => replacement)
  return insertMissingCell(worksheetXml, address, replacement)
}

function readXmlAttribute(xml: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(xml)?.[1]
}

function toA1Address(row: number, column: number): string {
  let letters = ''
  let remaining = column
  do {
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26) - 1
  } while (remaining >= 0)
  return `${letters}${row + 1}`
}

/// Mirrors the gateway's own cell insertion: a missing row or cell lands in
/// ascending order so Excel accepts the sheetData.
function insertMissingCell(worksheetXml: string, address: string, cellXml: string): string {
  const rowNumber = Number(/[1-9][0-9]*$/.exec(address)?.[0])
  if (!Number.isFinite(rowNumber)) throw new CellImageSaveError(`Invalid cell address: ${address}`)
  const targetColumn = parseColumn(address)
  const rowPattern = new RegExp(`<row\\b([^>]*?\\br="${rowNumber}"[^>]*?)(/>|>([\\s\\S]*?)</row>)`)
  const rowMatch = rowPattern.exec(worksheetXml)
  if (rowMatch) {
    const attributes = rowMatch[1] ?? ''
    const body = rowMatch[2] === '/>' ? '' : (rowMatch[3] ?? '')
    return worksheetXml.replace(
      rowPattern,
      () => `<row${attributes}>${insertCellInColumnOrder(body, cellXml, targetColumn)}</row>`,
    )
  }
  const newRow = `<row r="${rowNumber}">${cellXml}</row>`
  const rowStartPattern = /<row\b[^>]*?\br="([1-9][0-9]*)"/g
  let match: RegExpExecArray | null
  while ((match = rowStartPattern.exec(worksheetXml)) !== null) {
    if (Number(match[1]) > rowNumber) {
      return worksheetXml.slice(0, match.index) + newRow + worksheetXml.slice(match.index)
    }
  }
  if (worksheetXml.includes('</sheetData>')) {
    return worksheetXml.replace('</sheetData>', () => `${newRow}</sheetData>`)
  }
  const emptySheetData = /<sheetData\s*\/>/
  if (emptySheetData.test(worksheetXml)) {
    return worksheetXml.replace(emptySheetData, () => `<sheetData>${newRow}</sheetData>`)
  }
  throw new CellImageSaveError('Worksheet has no sheetData element.')
}

function insertCellInColumnOrder(rowBody: string, cellXml: string, targetColumn: number): string {
  const siblingPattern = /<c\b[^>]*?\br="([A-Z]{1,3})[1-9][0-9]*"/g
  let match: RegExpExecArray | null
  while ((match = siblingPattern.exec(rowBody)) !== null) {
    if (parseColumn(match[1] ?? '') > targetColumn) {
      return rowBody.slice(0, match.index) + cellXml + rowBody.slice(match.index)
    }
  }
  return rowBody + cellXml
}

function parseColumn(address: string): number {
  const letters = /^[A-Z]+/.exec(address)?.[0] ?? ''
  let column = 0
  for (const letter of letters) column = column * 26 + (letter.charCodeAt(0) - 64)
  return column - 1
}
