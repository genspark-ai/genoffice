/// Declarative defined-names save: rewrites workbook.xml's `<definedNames>`
/// from the editor's model. Entries the editor never models — `_xlnm.*`
/// built-ins, hidden names, and names it failed to install (`preserveNames`)
/// — stay byte-verbatim.

import { withFutureFunctionMarkers } from './future-functions'

export class DefinedNameError extends Error {}

export interface DefinedNameEntry {
  readonly name: string
  readonly formula: string
  /// localSheetId (position in workbook sheet order) for sheet-scoped names.
  readonly sheetIndex?: number | undefined
}

export interface DefinedNamesState {
  readonly names: readonly DefinedNameEntry[]
  readonly preserveNames: readonly string[]
}

/// Excel name rules (simplified): starts with a letter, `_`, or `\`;
/// continues with letters, digits, `_`, `.`, or `\`; must not look like an
/// A1 or R1C1 cell reference. Letters are Unicode — Excel accepts CJK names,
/// and Create from Selection builds them from localized headers, so this
/// must accept everything definedNameFromLabel produces.
const NAME_PATTERN = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u
const CELL_REF_PATTERN = /^(?:[A-Za-z]{1,3}[0-9]+|[Rr][0-9]*[Cc][0-9]*)$/

export function applyDefinedNamesState(workbookXml: string, state: DefinedNamesState): string {
  const preserved = new Set(state.preserveNames)
  const preservedKeys = new Set<string>()
  for (const match of workbookXml.matchAll(/<definedName\b[^>]*>/g)) {
    const name = /\bname="([^"]*)"/.exec(match[0])?.[1]
    if (name === undefined || !preserved.has(unescapeXml(name))) continue
    const scope = /\blocalSheetId="(\d+)"/.exec(match[0])?.[1]
    preservedKeys.add(`${unescapeXml(name)}\u0000${scope === undefined ? -1 : Number(scope)}`)
  }
  const seen = new Set<string>()
  for (const entry of state.names) {
    validateName(entry.name)
    const key = `${entry.name}\u0000${entry.sheetIndex ?? -1}`
    if (preservedKeys.has(key)) {
      throw new DefinedNameError(
        `The name "${entry.name}" also exists in a form the editor cannot model — ` +
          'saving would duplicate it.',
      )
    }
    if (seen.has(key)) {
      throw new DefinedNameError(`The name "${entry.name}" is defined twice.`)
    }
    seen.add(key)
  }

  // Drop every modeled entry from the existing section, keeping built-ins,
  // hidden names, and preserve-listed ones in place.
  const xml = workbookXml.replace(
    /<definedName\b[^>]*>[\s\S]*?<\/definedName>|<definedName\b[^>]*\/>/g,
    (element) => {
      const name = /\bname="([^"]*)"/.exec(element)?.[1] ?? ''
      const unescaped = unescapeXml(name)
      const keep =
        unescaped.startsWith('_xlnm') ||
        /\bhidden="(?:1|true)"/.test(element) ||
        preserved.has(unescaped)
      return keep ? element : ''
    },
  )

  const additions = state.names
    .map(
      (entry) =>
        `<definedName name="${escapeXmlAttribute(entry.name)}"` +
        (entry.sheetIndex === undefined ? '' : ` localSheetId="${entry.sheetIndex}"`) +
        `>${escapeXmlText(withFutureFunctionMarkers(entry.formula.replace(/^=/, '')))}</definedName>`,
    )
    .join('')

  return appendDefinedNames(xml, additions)
}

/**
 * Append entries to the workbook's `<definedNames>` section, creating the
 * section where there is none.
 *
 * Shared with the redaction mirror (`applyRedactionNameMirror`), which writes
 * into the same section with a different idea of what an entry means. The
 * schema order is the fiddly part — `<definedNames>` follows `<sheets>`, but
 * only after `<functionGroups>` and `<externalReferences>` when those are
 * present — so it is done once here rather than twice.
 */
export function appendDefinedNames(workbookXml: string, additions: string): string {
  const section = /<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>|<definedNames\b[^>]*\/>/.exec(
    workbookXml,
  )
  if (section) {
    const inner = (section[1] ?? '') + additions
    const replacement = inner === '' ? '' : `<definedNames>${inner}</definedNames>`
    return (
      workbookXml.slice(0, section.index) +
      replacement +
      workbookXml.slice(section.index + section[0].length)
    )
  }
  if (additions === '') return workbookXml
  // Schema order: definedNames follows sheets (and functionGroups/externalReferences).
  const anchor = /<\/sheets>|<sheets\b[^>]*\/>/.exec(workbookXml)
  if (!anchor) throw new DefinedNameError('workbook.xml has no sheets element.')
  const groups = /<functionGroups\b[^>]*>[\s\S]*?<\/functionGroups>|<functionGroups\b[^>]*\/>/.exec(
    workbookXml,
  )
  const externals =
    /<externalReferences\b[^>]*>[\s\S]*?<\/externalReferences>|<externalReferences\b[^>]*\/>/.exec(
      workbookXml,
    )
  let at = externals ? externals.index + externals[0].length : anchor.index + anchor[0].length
  // functionGroups only counts when externalReferences is absent: the later
  // element is the one the section must follow
  if (!externals && groups) at = groups.index + groups[0].length
  return `${workbookXml.slice(0, at)}<definedNames>${additions}</definedNames>${workbookXml.slice(at)}`
}

/** Shared by table names, which follow the same rules (minus the `_xlnm` reservation). */
export function isValidDefinedName(name: string): boolean {
  return (
    name.length > 0 &&
    name.length <= 255 &&
    NAME_PATTERN.test(name) &&
    !CELL_REF_PATTERN.test(name) &&
    name.toLowerCase() !== 'true' &&
    name.toLowerCase() !== 'false'
  )
}

function validateName(name: string): void {
  if (!isValidDefinedName(name)) {
    throw new DefinedNameError(`"${name}" is not a valid defined name.`)
  }
  if (name.startsWith('_xlnm')) {
    throw new DefinedNameError('Names starting with "_xlnm" are reserved by Excel.')
  }
}

export function unescapeXml(input: string): string {
  return input
    .replaceAll('&quot;', '"')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&')
}

export function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}
