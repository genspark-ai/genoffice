// Modified for ForaOffice: one inventory for protected deletion intent.
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { hasDeletionRevision } from './doc-utils'

export interface StructureItem {
  id: string
  kind: string
  from: number
  to: number
}

/** IDs describe one live snapshot; the runtime binds them to its document epoch. */
export function inventoryStructures(
  doc: ProseMirrorNode,
  from = 0,
  to = doc.content.size,
): StructureItem[] {
  const items: StructureItem[] = []
  const semanticKeys = new Map<StructureItem, string>()
  const add = (kind: string, start: number, end: number, semanticKey?: string) => {
    const previous = items[items.length - 1]
    if (
      semanticKey &&
      previous?.kind === kind &&
      previous.to === start &&
      semanticKeys.get(previous) === semanticKey
    ) {
      previous.to = end
      previous.id = `${kind}:${previous.from}:${end}`
      return
    }
    const item = { id: `${kind}:${start}:${end}`, kind, from: start, to: end }
    items.push(item)
    if (semanticKey) semanticKeys.set(item, semanticKey)
  }
  doc.descendants((node, pos) => {
    if (hasDeletionRevision(node) || node.marks.some((mark) => mark.type.name === 'del'))
      return false
    if (node.attrs.sectionBoundary) {
      add('section', pos + node.nodeSize - 1, pos + node.nodeSize)
      if (node.type.name === 'docProtected') return false
    }
    if (node.type.name === 'docProtected') {
      add('protected', pos, pos + node.nodeSize)
      return false
    }
    if (node.type.name === 'docTable' && from <= pos && to >= pos + node.nodeSize)
      add('table', pos, pos + node.nodeSize)
    if (node.type.name === 'docInlineImage') add('image', pos, pos + node.nodeSize)
    else if (node.type.name === 'docInlineMath') add('math', pos, pos + node.nodeSize)
    else if (['docNoteRef', 'docRuby', 'docCellBoxes', 'docNestedTable'].includes(node.type.name)) {
      add(node.type.name, pos, pos + node.nodeSize)
      return false
    }
    for (const mark of node.marks) {
      if (['instrField', 'refField', 'ctrlCheckbox', 'docSym'].includes(mark.type.name))
        add(mark.type.name, pos, pos + node.nodeSize, JSON.stringify(mark.attrs))
    }
    return true
  })
  return items.filter((item) => item.from < to && item.to > from)
}

const TABLE_GEOMETRY = [
  'colWidthsPct',
  'widthPx',
  'widthPct',
  'cellMar',
  'cellSpacingTwips',
  'borders',
  'tblAlign',
  'tblFloat',
  'tblAutoFit',
  'tblFixedLayout',
  'indentTwips',
  'bidiVisual',
]
const ROW_GEOMETRY = ['heightTwips', 'heightRule', 'repeatHeader', 'cantSplit']
const CELL_GEOMETRY = [
  'colspan',
  'rowspan',
  'colwidth',
  'gridGap',
  'mergeGap',
  'borders',
  'cellMar',
  'textDirection',
  'noWrap',
  'vAlign',
]

function tableGeometry(node: ProseMirrorNode): unknown {
  const keys =
    node.type.name === 'docTable'
      ? TABLE_GEOMETRY
      : node.type.name === 'docTableRow'
        ? ROW_GEOMETRY
        : CELL_GEOMETRY
  const attrs = Object.fromEntries(keys.map((key) => [key, node.attrs[key] ?? null]))
  const children: unknown[] = []
  if (node.type.name === 'docTable' || node.type.name === 'docTableRow')
    node.forEach((child) => children.push(tableGeometry(child)))
  return { type: node.type.name, attrs, children }
}

/** A compatible inherited table keeps its geometry while ordinary cell text changes. */
export function preservedTableStructures(
  doc: ProseMirrorNode,
  from: number,
  to: number,
  replacement: readonly ProseMirrorNode[],
): string[] {
  const next = replacement.filter((node) => node.type.name === 'docTable')
  const kept: string[] = []
  let ordinal = 0
  doc.forEach((node, pos) => {
    if (pos < from || pos + node.nodeSize > to || node.type.name !== 'docTable') return
    const candidate = next[ordinal++]
    if (
      candidate &&
      JSON.stringify(tableGeometry(node)) === JSON.stringify(tableGeometry(candidate))
    )
      kept.push(`table:${pos}:${pos + node.nodeSize}`)
  })
  return kept
}

/** Count exact retained atoms, including opaque content inherited by unchanged cells. */
export function preservedInlineStructures(
  doc: ProseMirrorNode,
  from: number,
  to: number,
  replacement: readonly ProseMirrorNode[],
  intent?: unknown,
): string[] {
  const next = doc.type.create(null, replacement)
  const eligible = (item: StructureItem, source: ProseMirrorNode) =>
    ['math', 'docRuby', 'docNoteRef', 'docCellBoxes', 'docNestedTable'].includes(item.kind) ||
    (item.kind === 'image' && !!source.nodeAt(item.from)?.attrs.rule)
  const candidates = inventoryStructures(next).filter((item) => eligible(item, next))
  const requested = Array.isArray(intent) ? intent : []
  // Identical atoms have identical payloads. Match retained occurrences first;
  // the declared deletion chooses which occurrence is omitted when counts shrink.
  const originals = inventoryStructures(doc, from, to)
    .filter((item) => eligible(item, doc))
    .sort((a, b) => Number(requested.includes(a.id)) - Number(requested.includes(b.id)))
  const kept: string[] = []
  for (const item of originals) {
    const original = doc.nodeAt(item.from)
    const index = candidates.findIndex(
      (candidate) => candidate.kind === item.kind && original?.eq(next.nodeAt(candidate.from)!),
    )
    if (index >= 0) {
      candidates.splice(index, 1)
      kept.push(item.id)
    }
  }
  return kept
}

/** No omitted, duplicated, extra or unrelated object may stand in for actual deletion. */
export function validateStructureDeletion(
  doc: ProseMirrorNode,
  from: number,
  to: number,
  intent: unknown,
  preserved: readonly string[] = [],
): string | null {
  const affected = inventoryStructures(doc, from, to).filter((item) => !preserved.includes(item.id))
  if (affected.some((item) => item.from < from || item.to > to))
    return 'This range cuts through a protected structure. No changes were made. Query and target the entire structure for explicit deletion, or edit ordinary text outside it.'
  if (
    intent !== undefined &&
    (!Array.isArray(intent) || intent.some((id) => typeof id !== 'string'))
  )
    return 'deleteStructures must be the exact list of current structure references to delete.'
  const requested = (intent ?? []) as string[]
  const actual = affected.map((item) => item.id).sort()
  if (JSON.stringify([...requested].sort()) === JSON.stringify(actual)) return null
  return 'This edit would delete protected document structures without exact matching deleteStructures intent. No changes were made. Query the current range, preserve its objects, or declare the exact structures to delete.'
}
