import { randomUUID } from 'node:crypto'

import {
  registerContentTypeOverride,
  relativeTarget,
  relsPathFor,
  resolveRelTarget,
  type MutablePackage,
} from './xlsx-drawing-add'
import { escapeRegExp, renameSheetReferencesInChart, SheetEditError } from './xlsx-sheets'

/// Relationship kinds (last segment of the Type URI) whose target belongs to
/// one owner and is copied for a duplicated sheet. Anything else below the
/// sheet level (images, embedded packages, theme overrides) stays shared.
const SHEET_OWNED = new Set(['comments', 'vmlDrawing', 'threadedComment', 'drawing', 'table'])
const NESTED_OWNED = new Set([
  'chart',
  'chartEx',
  'chartStyle',
  'chartColorStyle',
  'chartUserShapes',
])

export interface ClonedSheetParts {
  /// The duplicate's relationships part; null when nothing is left to relate.
  readonly relsXml: string | null
  readonly droppedPrinterSettings: boolean
  /// Source table name → the copy's table name, for the duplicate's formulas.
  readonly tableNames: ReadonlyMap<string, string>
}

interface CloneContext {
  readonly pkg: MutablePackage
  readonly touchedEntries: Set<string>
  readonly sourceName: string
  readonly cloneName: string
  readonly taken: Set<string>
  readonly tableNamesTaken: Set<string>
  readonly tableNames: Map<string, string>
  /// New part path → content, written once every part is known: the comment
  /// ids of a thread are shared by two parts.
  readonly pending: Map<string, string>
  nextTableId: number
}

/// Gives a duplicated sheet its own copy of the parts its source owns
/// (comments, drawings with their charts, tables), as Excel's Move or Copy
/// does, and returns the relationships part pointing at the copies.
export async function cloneSheetParts(
  pkg: MutablePackage,
  sourcePath: string,
  clonePath: string,
  sourceName: string,
  cloneName: string,
  touchedEntries: Set<string>,
): Promise<ClonedSheetParts> {
  const sourceRelsPath = relsPathFor(sourcePath)
  if (!(await pkg.has(sourceRelsPath))) {
    return { relsXml: null, droppedPrinterSettings: false, tableNames: new Map() }
  }
  const paths = await pkg.paths()
  const tablePaths = paths.filter((path) => /^xl\/tables\/[^/]+\.xml$/.test(path))
  const tableXmls = await Promise.all(tablePaths.map((path) => pkg.readText(path)))
  const workbookXml = await pkg.readText('xl/workbook.xml')
  const context: CloneContext = {
    pkg,
    touchedEntries,
    sourceName,
    cloneName,
    taken: new Set([...paths, ...(pkg.removedPaths?.() ?? [])]),
    tableNamesTaken: new Set([
      ...tableXmls.flatMap((xml) =>
        ['name', 'displayName'].flatMap((attribute) => {
          const value = tableAttribute(xml, attribute)
          return value === undefined ? [] : [value.toLowerCase()]
        }),
      ),
      ...[...workbookXml.matchAll(/<definedName\b[^>]*\bname="([^"]+)"/g)].map((match) =>
        (match[1] ?? '').toLowerCase(),
      ),
    ]),
    tableNames: new Map(),
    pending: new Map(),
    nextTableId: Math.max(0, ...tableXmls.map((xml) => Number(tableAttribute(xml, 'id') ?? 0))) + 1,
  }

  let droppedPrinterSettings = false
  let relsXml = await pkg.readText(sourceRelsPath)
  for (const relationship of parseRelationshipTags(relsXml)) {
    // a sheet background picture is media, shared like any other image
    if (relationship.external || relationship.kind === 'hyperlink' || relationship.kind === 'image')
      continue
    if (relationship.kind === 'printerSettings') {
      droppedPrinterSettings = true
      relsXml = relsXml.replace(relationship.tag, '')
      continue
    }
    if (!SHEET_OWNED.has(relationship.kind)) {
      throw new SheetEditError(
        `Sheet "${sourceName}" carries PivotTables, slicers, form controls, or embedded ` +
          'objects — duplicating it is not supported yet.',
      )
    }
    const copyPath = await clonePart(
      context,
      resolveRelTarget(sourcePath, relationship.target),
      relationship.kind,
    )
    relsXml = relsXml.replace(
      relationship.tag,
      retarget(relationship.tag, relativeTarget(clonePath, copyPath)),
    )
  }

  const threadIds = new Map<string, string>()
  for (const content of context.pending.values()) {
    for (const match of content.matchAll(/<threadedComment\b[^>]*?\bid="(\{[^}"]+\})"/g)) {
      const id = match[1]
      if (id && !threadIds.has(id)) threadIds.set(id, `{${randomUUID().toUpperCase()}}`)
    }
  }
  for (const [path, content] of context.pending) {
    // table formulas wait for every table of the sheet to have its new name
    let text = content.replace(
      TABLE_FORMULA,
      (_full, open: string, _tag: string, body: string, close: string) =>
        `${open}${renameTablesInFormula(body, context.tableNames)}${close}`,
    )
    for (const [from, to] of threadIds) text = text.split(from).join(to)
    pkg.add(path, text)
  }

  return {
    relsXml: /<Relationship\b/.test(relsXml) ? relsXml : null,
    droppedPrinterSettings,
    tableNames: context.tableNames,
  }
}

/// Points the duplicate's formulas at its own tables: Excel stores structured
/// references with the table name, and the copy's tables are renamed.
export function renameTablesInWorksheet(
  worksheetXml: string,
  tableNames: ReadonlyMap<string, string>,
): string {
  if (tableNames.size === 0) return worksheetXml
  return worksheetXml.replace(
    /(<(f|formula[12]?)\b(?:[^>]*[^/>])?>)([\s\S]*?)(<\/\2>)/g,
    (_full, open: string, _tag: string, body: string, close: string) =>
      `${open}${renameTablesInFormula(body, tableNames)}${close}`,
  )
}

const TABLE_FORMULA = /(<(calculatedColumnFormula|totalsRowFormula)\b[^>]*>)([\s\S]*?)(<\/\2>)/g

// string literals, quoted sheet names and bracketed column specifiers are text, not table names
const FORMULA_TEXT =
  '"[^"]*"|&quot;[\\s\\S]*?&quot;|\'[^\']*\'|&apos;[\\s\\S]*?&apos;' +
  "|\\[(?:'.|[^\\[\\]']|\\[(?:'.|[^\\[\\]'])*\\])*\\]"

function renameTablesInFormula(formula: string, tableNames: ReadonlyMap<string, string>): string {
  const copies = new Map([...tableNames].map(([from, to]) => [from.toLowerCase(), to]))
  const names = [...tableNames.keys()].sort((a, b) => b.length - a.length).map(escapeRegExp)
  return formula.replace(
    new RegExp(`${FORMULA_TEXT}|(?<![\\w.\\\\!])(?:${names.join('|')})(?![\\w.\\\\(!])`, 'gi'),
    (match) => copies.get(match.toLowerCase()) ?? match,
  )
}

async function clonePart(context: CloneContext, path: string, kind: string): Promise<string> {
  const { pkg } = context
  const copyPath = allocatePath(context.taken, path)
  let content = await pkg.readText(path)
  const relsPath = relsPathFor(path)
  const hasRels = await pkg.has(relsPath)

  if (kind === 'table') {
    if (hasRels) {
      throw new SheetEditError(
        `Sheet "${context.sourceName}" carries a table loaded from a query — ` +
          'duplicating it is not supported yet.',
      )
    }
    content = renameClonedTable(context, content)
  } else if (kind === 'chart' || kind === 'chartEx') {
    content = renameSheetReferencesInChart(content, context.sourceName, context.cloneName)
  }

  if (hasRels) {
    let relsXml = await pkg.readText(relsPath)
    for (const relationship of parseRelationshipTags(relsXml)) {
      if (relationship.external || !NESTED_OWNED.has(relationship.kind)) continue
      const nestedCopy = await clonePart(
        context,
        resolveRelTarget(path, relationship.target),
        relationship.kind,
      )
      relsXml = relsXml.replace(
        relationship.tag,
        retarget(relationship.tag, relativeTarget(copyPath, nestedCopy)),
      )
    }
    pkg.add(relsPathFor(copyPath), relsXml)
  }

  const contentType = await overrideContentType(pkg, path)
  if (contentType !== undefined) {
    await registerContentTypeOverride(pkg, copyPath, contentType, context.touchedEntries)
  }
  context.pending.set(copyPath, content)
  return copyPath
}

function renameClonedTable(context: CloneContext, tableXml: string): string {
  const name = tableAttribute(tableXml, 'displayName') ?? tableAttribute(tableXml, 'name')
  if (name === undefined) return tableXml
  let copyName = name
  for (let suffix = 2; context.tableNamesTaken.has(copyName.toLowerCase()); suffix += 1) {
    copyName = `${name}_${suffix}`
  }
  context.tableNamesTaken.add(copyName.toLowerCase())
  context.tableNames.set(name, copyName)
  const id = context.nextTableId
  context.nextTableId += 1
  return tableXml.replace(/<table\b[^>]*>/, (tag) =>
    tag
      .replace(/(\sid=")[^"]*(")/, `$1${id}$2`)
      .replace(/(\sname=")[^"]*(")/, `$1${copyName}$2`)
      .replace(/(\sdisplayName=")[^"]*(")/, `$1${copyName}$2`),
  )
}

function tableAttribute(tableXml: string, attribute: string): string | undefined {
  const tag = /<table\b[^>]*>/.exec(tableXml)?.[0]
  return tag === undefined ? undefined : new RegExp(`\\s${attribute}="([^"]*)"`).exec(tag)?.[1]
}

/// The next free numbered sibling of a part: xl/comments1.xml → xl/comments4.xml.
function allocatePath(taken: Set<string>, path: string): string {
  const match = /^(.*?)(\d*)(\.[^./]+)$/.exec(path)
  const stem = match?.[1] ?? path
  const extension = match?.[3] ?? ''
  for (let number = 1; ; number += 1) {
    const candidate = `${stem}${number}${extension}`
    if (taken.has(candidate)) continue
    taken.add(candidate)
    return candidate
  }
}

async function overrideContentType(pkg: MutablePackage, path: string): Promise<string | undefined> {
  const xml = await pkg.readText('[Content_Types].xml')
  for (const match of xml.matchAll(/<Override\b[^>]*\/?>/g)) {
    if (!match[0].includes(`PartName="/${path}"`)) continue
    return /\bContentType="([^"]+)"/.exec(match[0])?.[1]
  }
  return undefined
}

interface RelationshipTag {
  readonly tag: string
  readonly kind: string
  readonly target: string
  readonly external: boolean
}

function parseRelationshipTags(relsXml: string): RelationshipTag[] {
  const tags: RelationshipTag[] = []
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*?\/>\s*/g)) {
    const tag = match[0]
    const type = /\bType="([^"]+)"/.exec(tag)?.[1] ?? ''
    tags.push({
      tag,
      kind: type.slice(type.lastIndexOf('/') + 1),
      target: decodeAttribute(/\bTarget="([^"]*)"/.exec(tag)?.[1] ?? ''),
      external: /\bTargetMode="External"/.test(tag),
    })
  }
  return tags
}

function retarget(tag: string, target: string): string {
  return tag.replace(/(\bTarget=")[^"]*(")/, (_full, open: string, close: string) => {
    return `${open}${target.replace(/&/g, '&amp;').replace(/"/g, '&quot;')}${close}`
  })
}

function decodeAttribute(value: string): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}
