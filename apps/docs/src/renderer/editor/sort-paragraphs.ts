import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { NodeSelection, type EditorState } from '@tiptap/pm/state'
import { CellSelection, TableMap, selectedRect } from '@tiptap/pm/tables'
import { TRACK_IGNORE, activeTracking, revisionDateNow, withRowRecord } from './revisions'

export type SortBy = 'text' | 'number' | 'date'

export interface SortKey {
  /** 0 is the whole paragraph; n is its nth field, or the table's nth column */
  field: number
  by: SortBy
  descending: boolean
}

export interface SortOptions {
  keys: readonly SortKey[]
  /** the first paragraph or row names the fields and stays where it is */
  header: boolean
  /** what parts a paragraph into fields; a table's fields are its cells */
  separator: string
  caseSensitive: boolean
}

export const MAX_SORT_KEYS = 3

/** what Sort would act on: the rows of the table at the selection, or the paragraphs it touches */
export type SortScope =
  { kind: 'table'; rows: string[][] } | { kind: 'paragraphs'; texts: string[] }

export type SortRefusal = 'needsParagraphs' | 'mergedCells'

/** the first number in the text; text without one sorts first, as in Word */
const numberOf = (text: string) => {
  const match = /-?\d+(?:[.,]\d+)?/.exec(text)
  return match ? Number(match[0].replace(',', '.')) : -Infinity
}

const dateOf = (text: string) => {
  const time = Date.parse(text.trim())
  return Number.isNaN(time) ? -Infinity : time
}

// not a subtraction: two paragraphs without a key are both -Infinity, and their difference is NaN
const order = (a: number, b: number) => (a === b ? 0 : a < b ? -1 : 1)

export function compareParagraphs(by: SortBy, a: string, b: string, caseSensitive = false): number {
  if (by === 'number') return order(numberOf(a), numberOf(b))
  if (by === 'date') return order(dateOf(a), dateOf(b))
  return a.localeCompare(b, undefined, { sensitivity: caseSensitive ? 'variant' : 'base' })
}

/** `[whole text, field 1, field 2, …]`: a key's field is its index */
export function sortFields(text: string, separator: string): string[] {
  return [text, ...(separator ? text.split(separator) : [text])]
}

/** the first key that tells the two apart decides; a tie on every key keeps their order */
export function compareByKeys(
  keys: readonly SortKey[],
  a: readonly string[],
  b: readonly string[],
  caseSensitive = false,
): number {
  for (const key of keys) {
    const result = compareParagraphs(key.by, a[key.field] ?? '', b[key.field] ?? '', caseSensitive)
    if (result !== 0) return key.descending ? -result : result
  }
  return 0
}

interface Block {
  node: PmNode
  pos: number
}

interface TableTarget {
  table: PmNode
  pos: number
  /** the rows to sort, header included: `[top, bottom)` */
  top: number
  bottom: number
}

function tableDepth($pos: { depth: number; node: (depth: number) => PmNode }): number {
  for (let depth = $pos.depth; depth > 0; depth--)
    if ($pos.node(depth).type.name === 'docTable') return depth
  return -1
}

/** the table holding the whole selection; the rows a cell selection spans, or all of them */
function tableTarget(state: EditorState): TableTarget | null {
  const { selection } = state
  if (selection instanceof NodeSelection) {
    if (selection.node.type.name !== 'docTable') return null
    return { table: selection.node, pos: selection.from, top: 0, bottom: selection.node.childCount }
  }
  const fromDepth = tableDepth(selection.$from)
  const toDepth = tableDepth(selection.$to)
  if (fromDepth < 0 || toDepth < 0) return null
  const pos = selection.$from.before(fromDepth)
  if (pos !== selection.$to.before(toDepth)) return null
  const table = selection.$from.node(fromDepth)
  if (selection instanceof CellSelection) {
    const rect = selectedRect(state)
    if (rect.bottom - rect.top >= 2) return { table, pos, top: rect.top, bottom: rect.bottom }
  }
  return { table, pos, top: 0, bottom: table.childCount }
}

/** every row of the range is one plain cell per grid column */
function unmerged({ table, top, bottom }: TableTarget): boolean {
  const { width } = TableMap.get(table)
  for (let index = top; index < bottom; index++) {
    const row = table.child(index)
    if (row.childCount !== width) return false
    for (let cell = 0; cell < row.childCount; cell++) {
      const { colspan, rowspan, mergeGap, gridGap } = row.child(cell).attrs
      if (colspan !== 1 || rowspan !== 1 || mergeGap || gridGap) return false
    }
  }
  return true
}

const cellTexts = (row: PmNode) => {
  const texts: string[] = []
  row.forEach((cell) => texts.push(cell.textBetween(0, cell.content.size, ' ').trim()))
  return texts
}

function tableRows({ table, top, bottom }: TableTarget): PmNode[] {
  const rows: PmNode[] = []
  for (let index = top; index < bottom; index++) rows.push(table.child(index))
  return rows
}

/** the paragraphs the selection touches, or all of the body's with nothing selected */
function paragraphBlocks(state: EditorState): Block[] | null {
  const { doc, selection } = state
  const whole = selection.empty
  const blocks: Block[] = []
  let sortable = true
  doc.forEach((node, pos) => {
    const inRange = whole || (pos < selection.to && pos + node.nodeSize > selection.from)
    if (!inRange) return
    if (!node.isTextblock) sortable = false
    blocks.push({ node, pos })
  })
  return sortable ? blocks : null
}

/**
 * What the dialog offers keys for. A refusal when the table at the selection
 * has merged cells in the rows to sort, or the range holds anything but
 * paragraphs (a table among them, a field block).
 */
export function sortScope(state: EditorState): SortScope | SortRefusal {
  const target = tableTarget(state)
  if (target) {
    if (!unmerged(target)) return 'mergedCells'
    return { kind: 'table', rows: tableRows(target).map(cellTexts) }
  }
  const blocks = paragraphBlocks(state)
  if (!blocks) return 'needsParagraphs'
  return { kind: 'paragraphs', texts: blocks.map((block) => block.node.textContent) }
}

function sorted<T>(items: readonly T[], fields: (item: T) => string[], options: SortOptions): T[] {
  const head = options.header ? items.slice(0, 1) : []
  const keyed = items.slice(head.length).map((item) => ({ item, fields: fields(item) }))
  keyed.sort((a, b) => compareByKeys(options.keys, a.fields, b.fields, options.caseSensitive))
  return [...head, ...keyed.map((entry) => entry.item)]
}

/**
 * A tracked sort: each row that changed place is struck where it was and
 * inserted where it went. A row that is itself a pending insertion or
 * deletion just moves.
 */
function trackedRows(before: readonly PmNode[], after: readonly PmNode[], author: string) {
  let prefix = 0
  while (prefix < before.length && before[prefix] === after[prefix]) prefix++
  let suffix = 0
  while (
    suffix < before.length - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  )
    suffix++
  const date = revisionDateNow()
  const record = (row: PmNode, kind: 'ins' | 'del') =>
    row.type.create(
      {
        ...row.attrs,
        rowRevision: { kind, author, date },
        rawTrPr: withRowRecord(row.attrs.rawTrPr, kind, author, date),
      },
      row.content,
      row.marks,
    )
  const end = before.length - suffix
  return [
    ...before.slice(0, prefix),
    ...before
      .slice(prefix, end)
      .filter((row) => !row.attrs.rowRevision)
      .map((row) => record(row, 'del')),
    ...after.slice(prefix, end).map((row) => (row.attrs.rowRevision ? row : record(row, 'ins'))),
    ...before.slice(end),
  ]
}

function sortTableRows(editor: Editor, target: TableTarget, options: SortOptions): boolean {
  const rows = tableRows(target)
  if (rows.length - (options.header ? 1 : 0) < 2) return false
  const next = sorted(rows, (row) => ['', ...cellTexts(row)], options)
  if (next.every((row, index) => row === rows[index])) return true
  const { state } = editor
  const tracking = activeTracking(state)
  const { table, pos } = target
  const all: PmNode[] = []
  table.forEach((row) => all.push(row))
  all.splice(
    target.top,
    rows.length,
    ...(tracking ? trackedRows(rows, next, tracking.author) : next),
  )
  const tr = state.tr.replaceWith(
    pos,
    pos + table.nodeSize,
    // the rows carry their own cells: the save writes the table anew instead of patching texts in place
    table.type.create({ ...table.attrs, originalStructure: null }, all, table.marks),
  )
  if (tracking) tr.setMeta(TRACK_IGNORE, true)
  editor.view.dispatch(tr.scrollIntoView())
  return true
}

function sortParagraphBlocks(editor: Editor, blocks: Block[], options: SortOptions): boolean {
  if (blocks.length - (options.header ? 1 : 0) < 2) return false
  const next = sorted(
    blocks.map((block) => block.node),
    (node) => sortFields(node.textContent, options.separator),
    options,
  )
  if (next.every((node, index) => node === blocks[index].node)) return true
  const from = blocks[0].pos
  const last = blocks[blocks.length - 1]
  editor.view.dispatch(
    editor.state.tr.replaceWith(from, last.pos + last.node.nodeSize, next).scrollIntoView(),
  )
  return true
}

/**
 * Sorts the rows of the table at the selection by its columns, or else the
 * paragraphs the selection touches (all of the body's with nothing selected).
 * Null when it did; otherwise why nothing was moved.
 */
export function sortSelection(editor: Editor, options: SortOptions): SortRefusal | null {
  const target = tableTarget(editor.state)
  if (target) {
    if (!unmerged(target)) return 'mergedCells'
    return sortTableRows(editor, target, options) ? null : 'needsParagraphs'
  }
  const blocks = paragraphBlocks(editor.state)
  return blocks && sortParagraphBlocks(editor, blocks, options) ? null : 'needsParagraphs'
}
