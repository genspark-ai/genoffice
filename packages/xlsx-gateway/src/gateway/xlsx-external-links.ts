//! Breaking external workbook links (genoffice#2109): turns every formula
//! that references a broken `[n]` link into the literal its cached value
//! already holds, removes the link's parts and references, and renumbers the
//! `[m]` indices of the surviving links (the index is the position in
//! `<externalReferences>`, so removing one shifts every later one).

import { removePartOverride, removeRelationshipById } from './xlsx-sheets'
import type { EntrySource, MutationPlan } from './xlsx-gateway'

/// A stored formula body references the link `[n]`: the quoted spelling
/// `'[1]Sheet Name'!A1` (quotes doubled inside) or the bare `[1]Sheet1!A1`
/// (sheet-name characters run to the `!`, any non-ASCII included). The same
/// two shapes the sidecar's reference scan matches.
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

/// `[m]` → the link's new 1-based index after the broken set is removed.
function renumber(index: number, broken: readonly number[]): number {
  return index - broken.filter((n) => n < index).length
}

/// Replaces every `[m]` reference of a surviving link with its new index.
/// Runs over the whole sheet part: `<f>` bodies are the only place the
/// spelling occurs (workbook-level names do not carry link indices).
function renumberFormulaRefs(xml: string, broken: readonly number[]): string {
  return xml
    .replace(QUOTED_REF, (match, index: string, tail: string) => {
      const n = Number(index)
      return broken.includes(n) ? match : `'[${renumber(n, broken)}]${tail}'`
    })
    .replace(BARE_REF, (match, index: string, tail: string) => {
      const n = Number(index)
      return broken.includes(n) ? match : `[${renumber(n, broken)}]${tail}`
    })
}

/// True when the formula body references one of the `broken` links.
function referencesBroken(body: string, broken: readonly number[]): boolean {
  for (const n of referencedLinks(body)) {
    if (broken.includes(n)) return true
  }
  return false
}

/// Turns a formula cell into the literal its cached `<v>` already holds:
/// numbers and booleans keep their spelling, a formula-string result becomes
/// an inline string, an error result keeps its `t="e"` form.
function literalCell(cellXml: string): string {
  const reference = /\br="([^"]+)"/.exec(cellXml)?.[1] ?? ''
  const value = /<v>([\s\S]*?)<\/v>/.exec(cellXml)?.[1] ?? ''
  const kind = /\bt="(\w+)"/.exec(cellXml)?.[1]
  if (kind === 'str') {
    // a formula's string result: a literal string is an inline string
    return `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${value}</t></is></c>`
  }
  if (kind === 'b' || kind === 'e') {
    return `<c r="${reference}" t="${kind}"><v>${value}</v></c>`
  }
  return `<c r="${reference}"><v>${value}</v></c>`
}

/// Rewrites one sheet part: broken-link formulas become literals, surviving
/// links renumber. Instances of shared formulas carry no body and follow
/// their anchor; a shared anchor referencing a broken link cannot be broken
/// in place (the instances would keep computing) and fails the whole break.
export function transformSheetForLinkBreak(xml: string, broken: readonly number[]): string {
  // shared anchors: si → the formula body inside <f>
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
  const cellPattern = /<c\b[^>]*>(?:(?!<\/c>)[\s\S])*<\/c>/g
  return xml.replace(cellPattern, (cellXml) => {
    const body = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(cellXml)?.[1]
    if (body === undefined || !referencesBroken(body, broken)) {
      // surviving formulas renumber in place
      return body === undefined ? cellXml : renumberFormulaRefs(cellXml, broken)
    }
    return renumberFormulaRefs(literalCell(cellXml), broken)
  })
}

/// Plans the break of the given 1-based link indices: broken formulas become
/// literals, the links' parts and references disappear, surviving `[m]`
/// indices renumber. The source is read; the returned plan is what an
/// assembler writes.
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
      throw new Error(`no external link ${n}: the workbook has ${ordered.length}`)
    }
  }
  const brokenEntries = broken.map((n) => ({
    n,
    rId: ordered[n - 1]!.rId,
  }))
  const relationshipFor = (rId: string) =>
    new RegExp(`<Relationship\\b[^>]*\\bId="${rId}"[^>]*/>`).exec(relsXml)?.[0] ?? ''
  const partOf = (rId: string): string | null => {
    const element = relationshipFor(rId)
    const type = /\bType="([^"]+)"/.exec(element)?.[1] ?? ''
    if (!type.endsWith('/externalLink')) return null
    const target = /\bTarget="([^"]+)"/.exec(element)?.[1] ?? ''
    return target.startsWith('/') ? target.slice(1) : `xl/${target}`
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
      new RegExp(`<externalReference\\b[^>]*\\bid="${entry.rId}"[^>]*/>`),
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
  replaced.set('xl/workbook.xml', nextWorkbookXml)
  replaced.set('xl/_rels/workbook.xml.rels', nextRelsXml)

  let contentTypesXml: string | null = null
  if (await source.has('[Content_Types].xml')) {
    contentTypesXml = await source.readText('[Content_Types].xml')
  }
  const touched = ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']
  for (const entry of brokenParts) {
    removedEntries.push(entry.part, `${entry.part.replace(/([^/]+)$/, '_rels/$1')}.rels`)
    if (contentTypesXml !== null) {
      contentTypesXml = removePartOverride(contentTypesXml, entry.part)
    }
  }
  if (contentTypesXml !== null) {
    replaced.set('[Content_Types].xml', contentTypesXml)
    touched.push('[Content_Types].xml')
  }

  for (const path of await source.paths()) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) continue
    const xml = await source.readText(path)
    const transformed = transformSheetForLinkBreak(xml, broken)
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
