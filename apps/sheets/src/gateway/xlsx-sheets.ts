import { FORMULA_REFERENCE_PATTERN, qualifierMatches } from './xlsx-structure'

/// Sheet-level workbook surgery: rename, add, and remove worksheets. Renames
/// rewrite every qualified reference (other sheets' formulas, defined names,
/// chart series, hyperlink anchors); removals fail closed when anything still
/// references the sheet. All operations rebuild workbook.xml's sheet list in
/// the caller-provided final tab order.

export class SheetEditError extends Error {}

export interface SheetEditPlan {
  /// original file name → new name, for sheets already in the package
  readonly renames: readonly { readonly sheetName: string; readonly newName: string }[]
  /// final names of new sheets; a sourceSheetName means the part is seeded by
  /// cloning that sheet's part instead of a blank template
  readonly additions: readonly {
    readonly name: string
    readonly sourceSheetName?: string | undefined
  }[]
  /// original file names of sheets to delete
  readonly removals: readonly string[]
  /// the complete final tab order, as final names
  readonly order: readonly string[]
  /// visibility toggles, keyed by original (pre-rename) or added name
  readonly hiddenChanges?: readonly { readonly sheetName: string; readonly hidden: boolean }[]
  /// True when the tab order differs from the file — calcChain sheet indexes
  /// go stale, so the save drops it for Excel to rebuild.
  readonly orderChanged?: boolean
}

export interface SheetAllocation {
  readonly name: string
  readonly path: string
  readonly sheetId: number
  readonly relationshipId: string
}

const INVALID_NAME_CHARACTERS = /[\\/?*[\]:]/

export function validateSheetName(name: string): void {
  if (name.length === 0 || name.length > 31) {
    throw new SheetEditError(`Sheet name "${name}" must be 1-31 characters long.`)
  }
  if (INVALID_NAME_CHARACTERS.test(name)) {
    throw new SheetEditError(
      `Sheet name "${name}" contains a forbidden character (\\ / ? * [ ] :).`,
    )
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new SheetEditError(`Sheet name "${name}" cannot start or end with an apostrophe.`)
  }
}

/// Quotes a sheet name for use as a formula qualifier when Excel requires it
/// (non-identifier characters, or a name that parses as a cell reference).
export function formatSheetQualifier(name: string): string {
  const needsQuoting =
    !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ||
    /^[A-Za-z]{1,3}[0-9]+$/.test(name) ||
    /^(?:TRUE|FALSE)$/i.test(name) ||
    /^R[0-9]*C[0-9]*$/i.test(name)
  return needsQuoting ? `'${name.replaceAll("'", "''")}'` : name
}

/// Rewrites references qualified with oldName to newName inside one formula.
/// String literals ("" escaping) are left untouched.
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  return formula
    .split('"')
    .map((segment, index) =>
      index % 2 === 1
        ? segment
        : segment.replace(
            FORMULA_REFERENCE_PATTERN,
            (full, lead: string, qualifier: string | undefined, token: string) => {
              if (qualifier === undefined || !qualifierMatches(qualifier, oldName)) return full
              return `${lead}${formatSheetQualifier(newName)}!${token}`
            },
          ),
    )
    .join('"')
}

/// True when the formula contains a reference qualified with sheetName.
export function formulaReferencesSheet(formula: string, sheetName: string): boolean {
  let found = false
  formula.split('"').forEach((segment, index) => {
    if (index % 2 === 1 || found) return
    segment.replace(
      FORMULA_REFERENCE_PATTERN,
      (full, _lead: string, qualifier: string | undefined) => {
        if (qualifier !== undefined && qualifierMatches(qualifier, sheetName)) found = true
        return full
      },
    )
  })
  return found
}

/// Rewrites qualified references in a worksheet part: formula bodies, CF/DV
/// rule formulas, and internal hyperlink anchors.
export function renameSheetReferencesInWorksheet(
  worksheetXml: string,
  oldName: string,
  newName: string,
): string {
  // Self-closing `<f .../>` (shared-formula followers) must not match, or
  // the XML up to the next `</f>` would be swallowed as a formula body.
  let xml = worksheetXml.replace(
    /<f\b([^>]*[^/>])?>([\s\S]*?)<\/f>/g,
    (_full, attributes: string | undefined, body: string) =>
      `<f${attributes ?? ''}>${escapeXmlText(renameSheetInFormula(decodeEntities(body), oldName, newName))}</f>`,
  )
  xml = xml.replace(
    /<(formula[12]?)>([\s\S]*?)<\/\1>/g,
    (_full, tag: string, body: string) =>
      `<${tag}>${escapeXmlText(renameSheetInFormula(decodeEntities(body), oldName, newName))}</${tag}>`,
  )
  return xml.replace(
    /(<hyperlink\b[^>]*\blocation=")([^"]+)(")/g,
    (full, prefix: string, location: string, suffix: string) => {
      const renamed = renameHyperlinkLocation(decodeAttribute(location), oldName, newName)
      if (renamed === null) return full
      return `${prefix}${escapeXmlAttribute(renamed)}${suffix}`
    },
  )
}

function renameHyperlinkLocation(
  location: string,
  oldName: string,
  newName: string,
): string | null {
  const match = /^('(?:[^']|'')+'|[^'!]+)!([\s\S]*)$/.exec(location)
  if (!match?.[1] || match[2] === undefined) return null
  if (!qualifierMatches(match[1], oldName)) return null
  return `${formatSheetQualifier(newName)}!${match[2]}`
}

/// Rewrites qualified references in chart series (`<c:f>` bodies).
export function renameSheetReferencesInChart(
  chartXml: string,
  oldName: string,
  newName: string,
): string {
  return chartXml.replace(
    /(<c:f>)([\s\S]*?)(<\/c:f>)/g,
    (_full, open: string, body: string, close: string) =>
      `${open}${escapeXmlText(renameSheetInFormula(decodeEntities(body), oldName, newName))}${close}`,
  )
}

/// Rewrites qualified references in workbook definedName bodies.
export function renameSheetReferencesInDefinedNames(
  workbookXml: string,
  oldName: string,
  newName: string,
): string {
  return workbookXml.replace(
    /(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g,
    (_full, open: string, body: string, close: string) =>
      `${open}${escapeXmlText(renameSheetInFormula(decodeEntities(body), oldName, newName))}${close}`,
  )
}

export function worksheetReferencesSheet(worksheetXml: string, sheetName: string): boolean {
  for (const pattern of [
    /<f\b(?:[^>]*[^/>])?>([\s\S]*?)<\/f>/g,
    /<(?:formula[12]?)>([\s\S]*?)<\/(?:formula[12]?)>/g,
  ]) {
    for (const match of worksheetXml.matchAll(pattern)) {
      if (formulaReferencesSheet(decodeEntities(match[1] ?? ''), sheetName)) return true
    }
  }
  return false
}

export function chartReferencesSheet(chartXml: string, sheetName: string): boolean {
  for (const match of chartXml.matchAll(/<c:f>([\s\S]*?)<\/c:f>/g)) {
    if (formulaReferencesSheet(decodeEntities(match[1] ?? ''), sheetName)) return true
  }
  return false
}

/// Prepares a source sheet's relationships part for its duplicate. Hyperlink
/// relationships clone verbatim (ids are part-scoped, targets are external or
/// workbook-internal anchors). Printer settings reference a part the copy
/// must not share, so they are dropped — the caller strips the matching
/// r:id attributes. Anything else (drawings, tables, comments, pivots) would
/// leave the clone pointing at parts that cannot be shared: fail closed.
export function prepareClonedSheetRels(
  relsXml: string,
  sourceName: string,
): { relsXml: string | null; droppedPrinterSettings: boolean } {
  let droppedPrinterSettings = false
  let result = relsXml
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*?\bType="([^"]+)"[^>]*?\/>\s*/g)) {
    if (/\/hyperlink$/.test(match[1] ?? '')) continue
    if (/\/printerSettings$/.test(match[1] ?? '')) {
      droppedPrinterSettings = true
      result = result.replace(match[0], '')
      continue
    }
    throw new SheetEditError(
      `Sheet "${sourceName}" carries charts, images, tables, or comments — ` +
        'duplicating it is not supported yet.',
    )
  }
  return {
    relsXml: /<Relationship\b/.test(result) ? result : null,
    droppedPrinterSettings,
  }
}

/// Drops pageSetup r:id attributes from a cloned worksheet whose
/// printerSettings relationship was not cloned; the copy falls back to
/// default print settings instead of a dangling reference.
export function stripPageSetupRelIds(worksheetXml: string): string {
  return worksheetXml.replace(
    /(<pageSetup\b[^>]*?)\s+r:id="[^"]*"/g,
    (_full, prefix: string) => prefix,
  )
}

/// A clone must not claim the selected-tab state of its source.
export function sanitizeClonedWorksheetXml(worksheetXml: string): string {
  return worksheetXml.replace(
    /(<sheetView\b[^>]*?)\s+tabSelected="[^"]*"/g,
    (_full, prefix: string) => prefix,
  )
}

/// Excel clones sheet-scoped defined names when copying a sheet; this save
/// path does not yet, and silently diverging (formulas on the copy resolving
/// to a workbook-level name or #NAME?) is worse than refusing.
export function assertNoSheetScopedDefinedNames(workbookXml: string, sourceName: string): void {
  const index = parseSheetElements(workbookXml).findIndex((sheet) => sheet.name === sourceName)
  if (index < 0) throw new SheetEditError(`Sheet "${sourceName}" was not found in the workbook.`)
  if (new RegExp(`<definedName\\b[^>]*?\\blocalSheetId="${index}"`).test(workbookXml)) {
    throw new SheetEditError(
      `Sheet "${sourceName}" has sheet-scoped defined names — duplicating it is ` +
        'not supported yet.',
    )
  }
}

/// Blank worksheet part for added sheets.
export function buildWorksheetPartXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    '<sheetData/></worksheet>'
  )
}

export interface SheetElement {
  readonly xml: string
  readonly name: string
  readonly hidden: boolean
  readonly relationshipId: string | undefined
}

export function parseSheetElements(workbookXml: string): SheetElement[] {
  const elements: SheetElement[] = []
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*?(?:\/>|>[\s\S]*?<\/sheet>)/g)) {
    const xml = match[0]
    const name = readAttribute(xml, 'name')
    if (name === undefined) continue
    const state = readAttribute(xml, 'state')
    elements.push({
      xml,
      name: decodeAttribute(name),
      hidden: state === 'hidden' || state === 'veryHidden',
      // The relationships namespace is conventionally bound to "r", but any
      // prefix is legal — fall back to whatever prefix the producer chose.
      relationshipId:
        readAttribute(xml, 'r:id') ?? /(?:^|\s)[A-Za-z_][\w.-]*:id="([^"]*)"/.exec(xml)?.[1],
    })
  }
  return elements
}

export function maxSheetIdInWorkbook(workbookXml: string): number {
  let max = 0
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*?\bsheetId="([0-9]+)"/g)) {
    max = Math.max(max, Number(match[1]))
  }
  return max
}

export function maxRelationshipId(relationshipsXml: string): number {
  let max = 0
  for (const match of relationshipsXml.matchAll(/\bId="rId([0-9]+)"/g)) {
    max = Math.max(max, Number(match[1]))
  }
  return max
}

/// Rebuilds workbook.xml for the plan: sheet elements renamed / added /
/// removed and re-assembled in the final order, definedName localSheetId
/// scopes renumbered (or dropped with their sheet), and the active tab
/// clamped to a surviving visible sheet.
export function applySheetPlanToWorkbookXml(
  workbookXml: string,
  plan: SheetEditPlan,
  additions: readonly SheetAllocation[],
): string {
  const elements = parseSheetElements(workbookXml)
  const originalNames = elements.map((element) => element.name)
  const renameByOriginal = new Map(plan.renames.map((rename) => [rename.sheetName, rename.newName]))
  const removedSet = new Set(plan.removals)

  const finalElements = new Map<string, { xml: string; hidden: boolean }>()
  for (const element of elements) {
    if (removedSet.has(element.name)) continue
    const newName = renameByOriginal.get(element.name)
    const xml =
      newName === undefined
        ? element.xml
        : element.xml.replace(/\bname="[^"]*"/, () => `name="${escapeXmlAttribute(newName)}"`)
    finalElements.set(newName ?? element.name, { xml, hidden: element.hidden })
  }
  for (const addition of additions) {
    finalElements.set(addition.name, {
      xml: `<sheet name="${escapeXmlAttribute(addition.name)}" sheetId="${addition.sheetId}" r:id="${addition.relationshipId}"/>`,
      hidden: false,
    })
  }
  for (const change of plan.hiddenChanges ?? []) {
    const finalName = renameByOriginal.get(change.sheetName) ?? change.sheetName
    const element = finalElements.get(finalName)
    if (!element) {
      throw new SheetEditError(`Sheet "${change.sheetName}" was not found for a visibility change.`)
    }
    element.hidden = change.hidden
    element.xml = setSheetStateAttribute(element.xml, change.hidden)
  }

  if (
    plan.order.length !== finalElements.size ||
    plan.order.some((name) => !finalElements.has(name))
  ) {
    throw new SheetEditError('The sheet order does not match the final sheet set.')
  }
  if (plan.order.every((name) => finalElements.get(name)?.hidden)) {
    throw new SheetEditError('A workbook needs at least one visible sheet.')
  }

  const sheetsInner = plan.order.map((name) => finalElements.get(name)?.xml ?? '').join('')
  let result = workbookXml.replace(
    /<sheets>[\s\S]*?<\/sheets>/,
    () => `<sheets>${sheetsInner}</sheets>`,
  )

  // Defined-name scopes are positional (localSheetId = tab index).
  const oldIndexToNew = (oldIndex: number): number | null => {
    const originalName = originalNames[oldIndex]
    if (originalName === undefined || removedSet.has(originalName)) return null
    return plan.order.indexOf(renameByOriginal.get(originalName) ?? originalName)
  }
  result = result.replace(
    /<definedName\b[^>]*?\blocalSheetId="([0-9]+)"[^>]*>[\s\S]*?<\/definedName>/g,
    (full, localId: string) => {
      const mapped = oldIndexToNew(Number(localId))
      if (mapped === null) return ''
      return full.replace(/\blocalSheetId="[0-9]+"/, () => `localSheetId="${mapped}"`)
    },
  )
  result = result.replace(/<definedNames>\s*<\/definedNames>/, '')

  return result.replace(
    /(<workbookView\b[^>]*?\bactiveTab=")([0-9]+)(")/,
    (_full, prefix: string, activeTab: string, suffix: string) => {
      const mapped = oldIndexToNew(Number(activeTab))
      const fallback = plan.order.findIndex((name) => !finalElements.get(name)?.hidden)
      return `${prefix}${mapped ?? Math.max(fallback, 0)}${suffix}`
    },
  )
}

/// Toggles a `<sheet>` element's state attribute. Unhiding also clears
/// veryHidden — the only way a user reaches such a sheet is on purpose.
function setSheetStateAttribute(sheetXml: string, hidden: boolean): string {
  const withoutState = sheetXml.replace(/\s+state="[^"]*"/, '')
  if (!hidden) return withoutState
  return withoutState.replace(/^<sheet\b/, '<sheet state="hidden"')
}

export function addWorksheetOverride(contentTypesXml: string, partPath: string): string {
  const override = `<Override PartName="/${partPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
  if (!contentTypesXml.includes('</Types>')) {
    throw new SheetEditError('The workbook has no [Content_Types].xml Types element.')
  }
  return contentTypesXml.replace('</Types>', () => `${override}</Types>`)
}

/// Removes the [Content_Types].xml Override for any part path (worksheets,
/// drawings, charts, comments, tables). Parts covered by a Default extension
/// mapping have no Override; the call is a no-op for them.
export function removePartOverride(contentTypesXml: string, partPath: string): string {
  return contentTypesXml.replace(
    new RegExp(`<Override\\b[^>]*PartName="/${escapeRegExp(partPath)}"[^>]*/>\\s*`),
    '',
  )
}

export function addWorksheetRelationship(
  relationshipsXml: string,
  relationshipId: string,
  target: string,
): string {
  const relationship = `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${target}"/>`
  if (!relationshipsXml.includes('</Relationships>')) {
    throw new SheetEditError('The workbook has no relationships part.')
  }
  return relationshipsXml.replace('</Relationships>', () => `${relationship}</Relationships>`)
}

export function removeRelationshipById(relationshipsXml: string, relationshipId: string): string {
  return relationshipsXml.replace(
    new RegExp(`<Relationship\\b[^>]*\\bId="${escapeRegExp(relationshipId)}"[^>]*/>\\s*`),
    '',
  )
}

/// Relationship types whose targets never block a sheet removal: hyperlink
/// targets are external (or workbook-internal anchors that die with the
/// sheet's part), and printer settings survive as harmless orphans.
const IGNORABLE_SHEET_RELATIONSHIP_TYPES = /\/(?:hyperlink|printerSettings)$/

/// Relationship types whose targets are owned by the sheet: drawings (with
/// their images and charts), legacy VML, comments, threaded comments, form
/// control properties, tables, and a sheet-background image are anchored on
/// exactly one sheet, so they are cascade-deleted with it.
const OWNED_SHEET_RELATIONSHIP_TYPES =
  /\/(?:drawing|vmlDrawing|comments|threadedComment|ctrlProp|table|image)$/

export interface ParsedRelationship {
  readonly id: string | undefined
  readonly type: string
  readonly target: string
  readonly external: boolean
}

export function parseRelationships(relsXml: string): ParsedRelationship[] {
  const entries: ParsedRelationship[] = []
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*?\/>/g)) {
    const element = match[0]
    const type = readAttribute(element, 'Type')
    const target = readAttribute(element, 'Target')
    if (type === undefined || target === undefined) continue
    entries.push({
      id: readAttribute(element, 'Id'),
      type: decodeAttribute(type),
      target: decodeAttribute(target),
      external: readAttribute(element, 'TargetMode') === 'External',
    })
  }
  return entries
}

/// Inverse of xlsx-drawing-add's relsPathFor; the package-level "_rels/.rels"
/// maps to "".
export function partPathForRels(relsPath: string): string {
  return relsPath.replace(/(^|\/)_rels\/([^/]*)\.rels$/, '$1$2')
}

/// Returns the (unresolved) targets of a removed sheet's owned satellite
/// parts, for the caller to cascade-delete. Relationships that never block a
/// removal are skipped; anything else — pivot tables, slicers, OLE objects —
/// fails closed rather than leave dangling references for Excel to repair.
export function classifyRemovedSheetRels(relsXml: string, sheetName: string): string[] {
  const owned: string[] = []
  for (const entry of parseRelationships(relsXml)) {
    if (entry.external || IGNORABLE_SHEET_RELATIONSHIP_TYPES.test(entry.type)) continue
    if (OWNED_SHEET_RELATIONSHIP_TYPES.test(entry.type)) {
      owned.push(entry.target)
      continue
    }
    const kind = /\/pivotTable$/.test(entry.type)
      ? 'a pivot table'
      : `a part this build cannot delete safely (${entry.type})`
    throw new SheetEditError(
      `Sheet "${sheetName}" carries ${kind} — deleting it is not supported yet.`,
    )
  }
  return owned
}

/// displayName of the table defined in a table part — the token surviving
/// formulas would use in structured references (Table1[Amount]).
export function tableDisplayName(tableXml: string): string | undefined {
  const openTag = /<table\b[^>]*>/.exec(tableXml)?.[0]
  if (!openTag) return undefined
  const name = readAttribute(openTag, 'displayName') ?? readAttribute(openTag, 'name')
  return name === undefined ? undefined : decodeAttribute(name)
}

/// True when a defined name that survives the removal mentions the token
/// (a structured-reference prefix like "DecoTable["). Names scoped to a
/// removed sheet die with it and never block.
export function definedNamesUseToken(
  workbookXml: string,
  token: string,
  removedLocalIds: ReadonlySet<number>,
): boolean {
  for (const match of workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const localId = readAttribute(`<definedName${match[1] ?? ''}/>`, 'localSheetId')
    if (localId !== undefined && removedLocalIds.has(Number(localId))) continue
    if (decodeEntities(match[2] ?? '').includes(token)) return true
  }
  return false
}

/// DefinedNames scoped to a removed sheet are dropped with it, so their
/// bodies referencing the sheet must not block the removal.
export function definedNamesReferenceSheet(
  workbookXml: string,
  sheetName: string,
  removedLocalIds: ReadonlySet<number>,
): boolean {
  for (const match of workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const localId = readAttribute(`<definedName${match[1] ?? ''}/>`, 'localSheetId')
    if (localId !== undefined && removedLocalIds.has(Number(localId))) continue
    if (formulaReferencesSheet(decodeEntities(match[2] ?? ''), sheetName)) return true
  }
  return false
}

function readAttribute(elementXml: string, name: string): string | undefined {
  return new RegExp(`(?:^|\\s)${escapeRegExp(name)}="([^"]*)"`).exec(elementXml)?.[1]
}

function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

const NAMED_ENTITIES: Record<string, string> = {
  quot: '"',
  apos: "'",
  lt: '<',
  gt: '>',
  amp: '&',
}

/// Single-pass decode of the XML named entities plus numeric character
/// references (&#dd; / &#xhh;) — producers may encode non-ASCII sheet names
/// as numeric references, which sequential replaceAll would miss.
function decodeEntities(input: string): string {
  return input.replace(
    /&(?:#x([0-9A-Fa-f]+)|#([0-9]+)|(quot|apos|lt|gt|amp));/g,
    (match, hex: string | undefined, dec: string | undefined, named: string | undefined) => {
      if (named !== undefined) return NAMED_ENTITIES[named] ?? match
      const code = hex !== undefined ? Number.parseInt(hex, 16) : Number(dec)
      return code <= 0x10ffff ? String.fromCodePoint(code) : match
    },
  )
}

function decodeAttribute(input: string): string {
  return decodeEntities(input)
}

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return escapeXmlText(input).replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}
