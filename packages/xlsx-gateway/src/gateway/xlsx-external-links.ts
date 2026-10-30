//! Breaking external workbook links (genoffice#2109): turns every formula
//! that references a broken [n] link into the literal its cached value
//! already holds, removes the link's parts and references, and renumbers the
//! [m] indices of the surviving links (the index is the position in
//! <externalReferences>, so removing one shifts every later one).
//!
//! References live in more places than the cell formulas: a defined name, a
//! conditional-formatting rule, a data-validation formula and a chart range can
//! all carry a [n]Sheet!ref too. Those are renumbered with the surviving links;
//! a reference to a link being broken cannot become a literal in place, so the
//! whole break is refused explicitly rather than silently leaving a dangling
//! [n] behind.

import { removePartOverride, removeRelationshipById } from './xlsx-sheets'
import type { EntrySource, MutationPlan } from './xlsx-gateway'

/// A stored formula body references the link [n]: the quoted spelling
/// '[1]Sheet Name'!A1 (quotes doubled inside) or the bare [1]Sheet1!A1
/// (sheet-name characters run to the !, any non-ASCII included). The same two
/// shapes the sidecar's reference scan matches.
const QUOTED_REF = /'\[(\d+)\]((?:[^']|'')*)'/g
const BARE_REF = /\[(\d+)\]([^\s!(),;"'^&*<>=+\-*/[\]]+)/g

/// Indices of the links a formula body references.
function referencedLinks(body: string): Set<number> {
  const found = new Set<number>()
  for (const re of [QUOTED_REF, BARE_REF]) {
    re.lastIndex = 0
    for (const match of body.matchAll(re)) {
      found.add(Number(match[1]))
    }
  }
  return found
}

/// [m] -> the link's new 1-based index after the broken set is removed.
function renumber(index: number, broken: readonly number[]): number {
  return index - broken.filter((n) => n < index).length
}

/// Replaces every [m] reference of a surviving link with its new index.
/// Broken references are left for the caller to convert (cells) or refuse.
function renumberFormulaRefs(xml: string, broken: readonly number[]): string {
  return xml
    .replace(QUOTED_REF, (match, index: string, tail: string) => {
      const n = Number(index)
      return broken.includes(n) ? match : "'[" + renumber(n, broken) + ']' + tail + "'"
    })
    .replace(BARE_REF, (match, index: string, tail: string) => {
      const n = Number(index)
      return broken.includes(n) ? match : '[' + renumber(n, broken) + ']' + tail
    })
}

/// True when the formula body references one of the broken links.
function referencesBroken(body: string, broken: readonly number[]): boolean {
  for (const n of referencedLinks(body)) {
    if (broken.includes(n)) return true
  }
  return false
}

/// An element body that can carry a link reference: a cell formula, a
/// conditional-formatting rule formula, or a data-validation formula.
const SHEET_FORMULA_BODY = /<(f|formula|formula1|formula2)\b([^>]*)>([\s\S]*?)<\/\1>/g

/// Renumbers the surviving link references of the sheet's formula elements.
function renumberSheetFormulaBodies(xml: string, broken: readonly number[]): string {
  return xml.replace(SHEET_FORMULA_BODY, (all, tag: string, attributes: string, body: string) => {
    if (referencesBroken(body, broken)) return all
    const next = renumberFormulaRefs(body, broken)
    return next === body ? all : '<' + tag + attributes + '>' + next + '</' + tag + '>'
  })
}

/// Refuses the break when a reference to a broken link sits somewhere the
/// literal conversion cannot rewrite: a conditional-format rule, a validation,
/// an extension list, or anything else a future producer invents.
function assertNoBrokenSheetRefs(xml: string, broken: readonly number[]): void {
  for (const match of xml.matchAll(SHEET_FORMULA_BODY)) {
    if (!referencesBroken(match[3]!, broken)) continue
    const tag = match[1]!
    const where =
      tag === 'formula'
        ? 'a conditional-formatting rule'
        : tag === 'f'
          ? 'a cell formula'
          : 'a data-validation rule'
    throw new Error(
      'the sheet has ' +
        where +
        ' that references the external link to break; remove or rewrite that reference first',
    )
  }
  // Blank out the places already checked and the cell text a reader sees, so
  // only a reference in some other carrier is left to find.
  const rest = xml
    .replace(/<is>[\s\S]*?<\/is>/g, '<is/>')
    .replace(/<v\b[^>]*>[\s\S]*?<\/v>/g, '')
    .replace(/<t\b[^>]*>[\s\S]*?<\/t>/g, '')
    .replace(SHEET_FORMULA_BODY, '')
  if (referencesBroken(rest, broken)) {
    throw new Error(
      'the sheet references the external link to break outside cell formulas, ' +
        'conditional formatting and data validation; remove or rewrite that reference first',
    )
  }
}

/// The opening tag's attributes a literal keeps: everything except the cached
/// value's type, which the literal decides for itself, and the dynamic-array
/// mark, which a plain value must not claim.
function literalAttributes(cellXml: string): string {
  return (/^<c\b([^>]*)>/.exec(cellXml)?.[1] ?? '')
    .replace(/\s+t="[^"]*"/g, '')
    .replace(/\s+cm="[^"]*"/g, '')
}

/// Turns a formula cell into the literal its cached <v> already holds: numbers
/// and booleans keep their spelling, a formula-string result becomes an inline
/// string, an error result keeps its type. The cell's own attributes (style,
/// value metadata) survive; only the cache type goes. A formula with no usable
/// cache cannot become a literal (an empty <v> is how Excel repairs its file),
/// so it becomes the #REF! the broken link now means.
function literalCell(cellXml: string): string {
  const attributes = literalAttributes(cellXml)
  const kind = /\bt="(\w+)"/.exec(cellXml)?.[1]
  const cached = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(cellXml)
  if (cached === null) return '<c' + attributes + ' t="e"><v>#REF!</v></c>'
  const value = cached[1] ?? ''
  if (kind === 'str' || kind === 's' || kind === 'inlineStr') {
    // a formula's string result: a literal string is an inline string
    return (
      '<c' + attributes + ' t="inlineStr"><is><t xml:space="preserve">' + value + '</t></is></c>'
    )
  }
  if (kind === 'b' || kind === 'e') {
    return '<c' + attributes + ' t="' + kind + '"><v>' + value + '</v></c>'
  }
  // no type is a number, but an empty value is a producer artifact, not a value
  if (value === '') return '<c' + attributes + ' t="e"><v>#REF!</v></c>'
  return '<c' + attributes + '><v>' + value + '</v></c>'
}

/// Rewrites one sheet part: broken-link formulas become literals, surviving
/// links renumber. Instances of shared formulas carry no body and follow
/// their anchor; a shared anchor referencing a broken link cannot be broken
/// in place (the instances would keep computing) and fails the whole break.
export function transformSheetForLinkBreak(xml: string, broken: readonly number[]): string {
  // shared anchors: si -> the formula body inside <f>
  const sharedBodies = new Map<string, string>()
  for (const match of xml.matchAll(/<f t="shared"[^>]*\bsi="(\d+)"[^>]*>([\s\S]*?)<\/f>/g)) {
    sharedBodies.set(match[1]!, match[2]!)
  }
  for (const match of xml.matchAll(/<f t="shared"[^>]*\bsi="(\d+)"[^>]*\/>/g)) {
    const body = sharedBodies.get(match[1]!)
    if (body !== undefined && referencesBroken(body, broken)) {
      throw new Error(
        'a shared formula references the link to break; re-enter that cell ' +
          'as a literal first (edit the formula, then break the link)',
      )
    }
  }
  // A cell is either self-closing (<c r="A1" s="2"/>, a styled empty cell
  // Excel writes by the hundred) or has a body. The optional-/> opening keeps
  // the former from swallowing the next cell, whose </c> the old pattern
  // would run to.
  const cellPattern = /<c\b[^>]*?(?:\/>|>(?:(?!<\/c>)[\s\S])*<\/c>)/g
  const transformed = xml.replace(cellPattern, (cellXml) => {
    if (cellXml.endsWith('/>')) return cellXml
    const body = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(cellXml)?.[1]
    if (body === undefined || !referencesBroken(body, broken)) return cellXml
    return literalCell(cellXml)
  })
  // The refusal check reads the file-spelled indices: after the renumbering a
  // surviving link may legitimately take a number a broken one used to have.
  assertNoBrokenSheetRefs(transformed, broken)
  return renumberSheetFormulaBodies(transformed, broken)
}

/// Rewrites the workbook part: a defined-name body can reference links too, so
/// the surviving ones renumber there. A name that references a link being
/// broken has no literal form, so the break is refused.
export function transformWorkbookForLinkBreak(
  workbookXml: string,
  broken: readonly number[],
): string {
  return workbookXml.replace(
    /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g,
    (all, attributes: string, body: string) => {
      if (referencesBroken(body, broken)) {
        throw new Error(
          'a defined name references the external link to break; rewrite or ' +
            'remove that name first',
        )
      }
      const next = renumberFormulaRefs(body, broken)
      return next === body ? all : '<definedName' + attributes + '>' + next + '</definedName>'
    },
  )
}

/// Rewrites a chart part: a chart's <c:f> range can reference a link, so the
/// surviving ones renumber. A range that references a broken link breaks the
/// chart, so the break is refused instead of writing a dangling reference.
export function transformChartForLinkBreak(chartXml: string, broken: readonly number[]): string {
  return chartXml.replace(/<c:f>([\s\S]*?)<\/c:f>/g, (all, body: string) => {
    if (referencesBroken(body, broken)) {
      throw new Error(
        'a chart range references the external link to break; remove the chart ' +
          'or rewrite its data range first',
      )
    }
    const next = renumberFormulaRefs(body, broken)
    return next === body ? all : '<c:f>' + next + '</c:f>'
  })
}

/// Plans the break of the given 1-based link indices: broken formulas become
/// literals, the links' parts and references disappear, surviving [m] indices
/// renumber. The source is read; the returned plan is what an assembler
/// writes. Every refusal happens here, before anything is written.
export async function planExternalLinkBreak(
  source: EntrySource,
  breakIndices: readonly number[],
): Promise<MutationPlan> {
  const broken = [...new Set(breakIndices)].sort((a, b) => a - b)
  if (broken.length === 0 || broken.some((n) => !Number.isInteger(n) || n < 1)) {
    throw new Error('external link indices must be positive integers')
  }
  const workbookXml = await source.readText('xl/workbook.xml')
  const relsXml = await source.readText('xl/_rels/workbook.xml.rels')
  // the n-th <externalReference> is link [n]
  const ordered: { n: number; rId: string }[] = []
  for (const match of workbookXml.matchAll(/<externalReference\b[^>]*\br:id="([^"]+)"/g)) {
    ordered.push({ n: ordered.length + 1, rId: match[1]! })
  }
  for (const n of broken) {
    if (n > ordered.length) {
      throw new Error('no external link ' + n + ': the workbook has ' + ordered.length)
    }
  }
  const brokenEntries = broken.map((n) => ({
    n,
    rId: ordered[n - 1]!.rId,
  }))
  const relationshipFor = (rId: string) =>
    new RegExp('<Relationship\\b[^>]*\\bId="' + rId + '"[^>]*/>').exec(relsXml)?.[0] ?? ''
  const partOf = (rId: string): string | null => {
    const element = relationshipFor(rId)
    const type = /\bType="([^"]+)"/.exec(element)?.[1] ?? ''
    if (!type.endsWith('/externalLink')) return null
    const target = /\bTarget="([^"]+)"/.exec(element)?.[1] ?? ''
    return target.startsWith('/') ? target.slice(1) : 'xl/' + target
  }
  const brokenParts = brokenEntries
    .map((entry) => ({ ...entry, part: partOf(entry.rId) }))
    .filter((entry): entry is { n: number; rId: string; part: string } => entry.part !== null)

  const replaced = new Map<string, string>()
  const removedEntries: string[] = []
  let nextWorkbookXml = workbookXml
  let nextRelsXml = relsXml
  for (const entry of brokenEntries) {
    // the element's only attribute is the relationship id (r:id)
    nextWorkbookXml = nextWorkbookXml.replace(
      new RegExp('<externalReference\\b[^>]*\\bid="' + entry.rId + '"[^>]*/>'),
      '',
    )
    nextRelsXml = removeRelationshipById(nextRelsXml, entry.rId)
  }
  if (!/<externalReference\b/.test(nextWorkbookXml)) {
    nextWorkbookXml = nextWorkbookXml.replace(
      /<externalReferences\b[^>]*>[\s\S]*?<\/externalReferences>/,
      '',
    )
  }
  // defined names address links through the same [n] indices, so a surviving
  // name has to follow the renumbering; a broken one has no literal form
  nextWorkbookXml = transformWorkbookForLinkBreak(nextWorkbookXml, broken)
  replaced.set('xl/workbook.xml', nextWorkbookXml)
  replaced.set('xl/_rels/workbook.xml.rels', nextRelsXml)

  let contentTypesXml: string | null = null
  if (await source.has('[Content_Types].xml')) {
    contentTypesXml = await source.readText('[Content_Types].xml')
  }
  const touched = ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']
  for (const entry of brokenParts) {
    removedEntries.push(entry.part, entry.part.replace(/([^/]+)$/, '_rels/$1') + '.rels')
    if (contentTypesXml !== null) {
      contentTypesXml = removePartOverride(contentTypesXml, entry.part)
    }
  }
  if (contentTypesXml !== null) {
    replaced.set('[Content_Types].xml', contentTypesXml)
    touched.push('[Content_Types].xml')
  }

  for (const path of await source.paths()) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(path)) continue
    const xml = await source.readText(path)
    const transformed = transformSheetForLinkBreak(xml, broken)
    if (transformed !== xml) {
      replaced.set(path, transformed)
      touched.push(path)
    }
  }
  // chart ranges reference cells (and thus links) outside the sheet part
  for (const path of await source.paths()) {
    if (!/^xl\/charts\/[^/]+\.xml$/.test(path)) continue
    const xml = await source.readText(path)
    const transformed = transformChartForLinkBreak(xml, broken)
    if (transformed !== xml) {
      replaced.set(path, transformed)
      touched.push(path)
    }
  }
  return {
    replaced,
    added: new Map(),
    addedBinary: new Map(),
    removedEntries,
    addedEntries: [],
    touchedEntries: touched,
  }
}
