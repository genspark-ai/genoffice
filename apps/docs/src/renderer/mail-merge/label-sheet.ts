import type { Editor } from '@tiptap/core'
import type { Node as PmDocNode } from '@tiptap/pm/model'
import type { PmNode } from '../editor/convert'
import { cellContentOf, emptyParagraph, labelSheetDoc, type MergePageLayout } from './layout'

/**
 * A labels main document, as Word lays one out: a table with a label in each
 * cell. The first label holds the layout; every other label opens with a NEXT
 * field, so each cell shows the following record, and Update Labels copies
 * the first label into the rest.
 */

const NEXT_RE = /^\s*NEXT\s*$/i

const isNext = (node: PmNode): boolean =>
  node.type === 'text' &&
  (node.marks ?? []).some(
    (m) => m.type === 'instrField' && NEXT_RE.test(String(m.attrs?.instr ?? '')),
  )

const hasNext = (node: PmNode): boolean => isNext(node) || (node.content ?? []).some(hasNext)

/** the cells that carry labels: as wide as the first one (gutter cells and gutter rows are not) */
function labelCells(table: PmNode): PmNode[] {
  const cells = (table.content ?? []).flatMap((row) => row.content ?? [])
  const width = (cell: PmNode) => (cell.attrs?.colwidth as number[] | null | undefined)?.[0] ?? null
  const first = cells.find((cell) => Number(cell.attrs?.colspan ?? 1) === 1)
  if (!first) return []
  const w = width(first)
  return cells.filter(
    (cell) =>
      Number(cell.attrs?.colspan ?? 1) === 1 &&
      (w === null || width(cell) === null || Math.abs((width(cell) as number) - w) <= 2),
  )
}

/** index of the body's label table: the first table with a NEXT field past its first label */
export function labelTableIndex(doc: PmNode): number {
  return (doc.content ?? []).findIndex(
    (block) => block.type === 'docTable' && labelCells(block).slice(1).some(hasNext),
  )
}

/** whether the body holds a label table; walks only tables that carry a NEXT field */
export function hasLabelSheet(doc: PmDocNode): boolean {
  let found = false
  doc.forEach((block) => {
    if (found || block.type.name !== 'docTable') return
    let next = false
    block.descendants((node) => {
      if (next) return false
      if (
        node.isText &&
        node.marks.some(
          (m) => m.type.name === 'instrField' && NEXT_RE.test(String(m.attrs.instr ?? '')),
        )
      )
        next = true
      return !next
    })
    if (next)
      found = labelCells(block.toJSON() as PmNode)
        .slice(1)
        .some(hasNext)
  })
  return found
}

/** the highest field id in use: each NEXT field needs one of its own to act on its own */
function maxFieldId(node: PmNode): number {
  let max = 0
  for (const m of node.marks ?? [])
    if (m.type === 'instrField') max = Math.max(max, Number(m.attrs?.fieldId) || 0)
  for (const child of node.content ?? []) max = Math.max(max, maxFieldId(child))
  return max
}

type FieldIds = { next: number }

/** the NEXT field a label opens with */
function nextField(ids: FieldIds): PmNode {
  return {
    type: 'text',
    text: ' ',
    marks: [
      { type: 'instrField', attrs: { instr: 'NEXT', fieldId: ids.next++, fieldPart: 'single' } },
    ],
  }
}

const withoutNext = (nodes: readonly PmNode[]): PmNode[] =>
  nodes
    .filter((node) => !isNext(node))
    .map((node) => (node.content ? { ...node, content: withoutNext(node.content) } : node))

/** the first label's blocks as another label carries them: after a NEXT field */
function followingLabel(first: readonly PmNode[], ids: FieldIds): PmNode[] {
  const blocks = withoutNext(first)
  const [head, ...rest] = blocks.length > 0 ? blocks : [emptyParagraph()]
  return [{ ...head, content: [nextField(ids), ...(head.content ?? [])] }, ...rest]
}

/** the table with every label after the first rewritten from the first */
function updatedLabelTable(table: PmNode, ids: FieldIds): PmNode {
  const labels = labelCells(table)
  const first = labels[0]
  if (!first) return table
  const others = new Set(labels.slice(1))
  return {
    ...table,
    content: (table.content ?? []).map((row) => ({
      ...row,
      content: (row.content ?? []).map((cell) =>
        others.has(cell) ? { ...cell, content: followingLabel(first.content ?? [], ids) } : cell,
      ),
    })),
  }
}

function replaceBlock(editor: Editor, index: number, json: PmNode): boolean {
  const { state } = editor
  let pos = 0
  for (let i = 0; i < index; i++) pos += state.doc.child(i).nodeSize
  const old = state.doc.child(index)
  let node: PmDocNode
  try {
    node = state.schema.nodeFromJSON(json)
    node.check()
  } catch {
    return false
  }
  editor.view.dispatch(state.tr.replaceWith(pos, pos + old.nodeSize, node))
  return true
}

/** Mailings > Update Labels: false when the document has no label table */
export function updateLabels(editor: Editor): boolean {
  const doc = editor.getJSON() as PmNode
  const index = labelTableIndex(doc)
  if (index < 0) return false
  return replaceBlock(
    editor,
    index,
    updatedLabelTable(doc.content![index], { next: maxFieldId(doc) + 1 }),
  )
}

/**
 * Turns the document into a sheet of labels on the layout's grid: what the
 * document held becomes the first label and the rest follow it. The caller
 * sets the paper.
 */
export function createLabelSheet(editor: Editor, layout: MergePageLayout): boolean {
  if (!layout.grid) return false
  const doc = editor.getJSON() as PmNode
  // a document that is a label sheet already is laid out again from its first label
  const sheetAt = labelTableIndex(doc)
  const label = sheetAt >= 0 ? labelCells(doc.content![sheetAt])[0] : undefined
  const first = label ? withoutNext(label.content ?? []) : cellContentOf(doc)
  const ids = { next: maxFieldId(doc) + 1 }
  const count = layout.grid.columns * layout.grid.rows
  const cells = Array.from({ length: count }, (_, i) =>
    i === 0 ? first : followingLabel(first, ids),
  )
  const sheet = labelSheetDoc(cells, layout)
  const { state } = editor
  let content
  try {
    content = state.schema.nodeFromJSON(sheet).content
  } catch {
    return false
  }
  editor.view.dispatch(state.tr.replaceWith(0, state.doc.content.size, content))
  return true
}

/** a cell the merge emptied (its only paragraph was a NEXT field) still needs a paragraph */
const filledRow = (row: PmNode): PmNode => ({
  ...row,
  content: (row.content ?? []).map((cell) =>
    (cell.content ?? []).length > 0 ? cell : { ...cell, content: [emptyParagraph()] },
  ),
})

/**
 * Merged copies of a labels main document as one sheet after another: the
 * later copies' label rows join the first copy's table, the block at `at`.
 */
export function joinLabelSheets(docs: readonly PmNode[], at: number): PmNode | null {
  const [head, ...rest] = docs
  const table = head?.content?.[at]
  if (!head || table?.type !== 'docTable') return null
  const rows = [
    ...(table.content ?? []),
    ...rest.flatMap((doc) => doc.content?.[at]?.content ?? []),
  ].map(filledRow)
  return {
    ...head,
    content: head.content!.map((block, i) => (i === at ? { ...table, content: rows } : block)),
  }
}
