/**
 * Word's Navigation pane model: the heading tree, the "heading and its content"
 * range (heading start to the next heading of the same or a higher level), search
 * hit counts per subtree, and the subtree move/delete/level edits as transactions.
 */
import type { Schema } from '@tiptap/pm/model'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import type { StyleInfo } from '@genoffice/docx-engine'
import type { HeadingRef } from './headings'
import { matchContext, type FindMatch } from './find'

export const MAX_HEADING_LEVEL = 9
export const NAV_CONTEXT_RADIUS = 40

export interface NavNode {
  index: number
  ref: HeadingRef
  parent: number
  children: NavNode[]
}

/** nearest earlier heading with a smaller level is the parent (skipped levels nest under it) */
export function buildHeadingTree(headings: HeadingRef[]): NavNode[] {
  const roots: NavNode[] = []
  const stack: NavNode[] = []
  headings.forEach((ref, index) => {
    while (stack.length && stack[stack.length - 1].ref.level >= ref.level) stack.pop()
    const parent = stack[stack.length - 1]
    const node: NavNode = { index, ref, parent: parent ? parent.index : -1, children: [] }
    if (parent) parent.children.push(node)
    else roots.push(node)
    stack.push(node)
  })
  return roots
}

/** end of the heading's own text block range: the next heading of any level */
export function ownEnd(headings: HeadingRef[], index: number, docSize: number): number {
  return index + 1 < headings.length ? headings[index + 1].pos : docSize
}

/** end of heading + content: the next heading of the same or a higher level */
export function subtreeEnd(headings: HeadingRef[], index: number, docSize: number): number {
  const level = headings[index].level
  for (let j = index + 1; j < headings.length; j++)
    if (headings[j].level <= level) return headings[j].pos
  return docSize
}

export function subtreeRange(
  headings: HeadingRef[],
  index: number,
  docSize: number,
): { from: number; to: number } {
  return { from: headings[index].pos, to: subtreeEnd(headings, index, docSize) }
}

/** index of the heading whose section contains `pos` (last heading at or before it), -1 when none */
export function headingAtPos(headings: HeadingRef[], pos: number): number {
  let lo = 0
  let hi = headings.length - 1
  let found = -1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    if (headings[mid].pos <= pos) {
      found = mid
      lo = mid + 1
    } else hi = mid - 1
  }
  return found
}

/** collapse state survives edits that move the heading: keyed by level + text, not position */
export function headingKey(ref: Pick<HeadingRef, 'level' | 'text'>): string {
  return `${ref.level}\u0000${ref.text}`
}

export interface NavRow {
  node: NavNode
  depth: number
  hasChildren: boolean
  collapsed: boolean
}

/** pre-order rows of the tree with collapsed subtrees and levels above `maxLevel` hidden */
export function visibleRows(
  tree: NavNode[],
  collapsed: ReadonlySet<string>,
  maxLevel: number,
): NavRow[] {
  const rows: NavRow[] = []
  const walk = (nodes: NavNode[], depth: number) => {
    for (const node of nodes) {
      if (node.ref.level > maxLevel) continue
      const children = node.children.filter((c) => c.ref.level <= maxLevel)
      const isCollapsed = children.length > 0 && collapsed.has(headingKey(node.ref))
      rows.push({ node, depth, hasChildren: children.length > 0, collapsed: isCollapsed })
      if (!isCollapsed) walk(node.children, depth + 1)
    }
  }
  walk(tree, 0)
  return rows
}

/** search hits per heading: `own` inside the heading's text block range, `subtree` including nested headings */
export function hitCounts(
  headings: HeadingRef[],
  hits: ReadonlyArray<{ from: number }>,
  docSize: number,
): { own: number[]; subtree: number[] } {
  const own = new Array<number>(headings.length).fill(0)
  const subtree = new Array<number>(headings.length).fill(0)
  if (headings.length === 0) return { own, subtree }
  const sorted = hits.map((h) => h.from).sort((a, b) => a - b)
  let k = 0
  for (let i = 0; i < headings.length; i++) {
    const end = ownEnd(headings, i, docSize)
    while (k < sorted.length && sorted[k] < headings[i].pos) k++
    while (k < sorted.length && sorted[k] < end) {
      own[i]++
      k++
    }
  }
  const stack: number[] = []
  for (let i = 0; i < headings.length; i++) {
    while (stack.length && headings[stack[stack.length - 1]].level >= headings[i].level) stack.pop()
    subtree[i] += own[i]
    for (const anc of stack) subtree[anc] += own[i]
    stack.push(i)
  }
  return { own, subtree }
}

export function resultContext(m: FindMatch): { before: string; hit: string; after: string } {
  return matchContext(m, NAV_CONTEXT_RADIUS)
}

/**
 * Move a heading with its content before another heading or after another heading's
 * subtree. Both boundaries are top-level block edges, so the cut slice is closed and
 * one delete + one insert on the same transaction is one undo step.
 */
export function moveHeadingSubtree(
  tr: Transaction,
  headings: HeadingRef[],
  from: number,
  target: number,
  where: 'before' | 'after',
): boolean {
  if (from === target) return false
  const docSize = tr.doc.content.size
  const src = subtreeRange(headings, from, docSize)
  const insertAt = where === 'before' ? headings[target].pos : subtreeEnd(headings, target, docSize)
  if (insertAt >= src.from && insertAt <= src.to) return false
  const slice = tr.doc.slice(src.from, src.to)
  tr.delete(src.from, src.to)
  const dest = insertAt > src.to ? insertAt - (src.to - src.from) : insertAt
  tr.insert(dest, slice.content)
  tr.setSelection(TextSelection.near(tr.doc.resolve(dest + 1), 1))
  return true
}

/** Word promotes/demotes the heading alone; style-based headings switch to the style of the new level */
export function setHeadingLevel(
  tr: Transaction,
  schema: Schema,
  pos: number,
  level: number,
  styles?: ReadonlyMap<string, Pick<StyleInfo, 'headingLevel' | 'headingOutlineOff'>>,
): boolean {
  const clamped = Math.min(Math.max(level, 1), MAX_HEADING_LEVEL)
  const node = tr.doc.nodeAt(pos)
  if (!node || !node.isTextblock) return false
  if (node.type.name === 'docHeading') {
    if (Number(node.attrs.level) === clamped) return false
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, level: clamped })
    return true
  }
  if (styles) {
    for (const [styleId, style] of styles) {
      if (style.headingLevel === clamped && !style.headingOutlineOff) {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, styleId })
        return true
      }
    }
  }
  const heading = schema.nodes.docHeading
  if (!heading) return false
  tr.setNodeMarkup(pos, heading, { ...node.attrs, styleId: null, level: clamped })
  return true
}

/** Word's New Heading Before / After: an empty heading of the same level, caret inside */
export function insertHeadingSibling(
  tr: Transaction,
  schema: Schema,
  headings: HeadingRef[],
  index: number,
  where: 'before' | 'after',
): boolean {
  const heading = schema.nodes.docHeading
  if (!heading) return false
  const docSize = tr.doc.content.size
  const at = where === 'before' ? headings[index].pos : subtreeEnd(headings, index, docSize)
  tr.insert(at, heading.create({ level: headings[index].level }))
  tr.setSelection(TextSelection.create(tr.doc, at + 1))
  return true
}

/** Delete heading + content; a document must keep one block */
export function deleteHeadingSubtree(
  tr: Transaction,
  schema: Schema,
  headings: HeadingRef[],
  index: number,
): boolean {
  const docSize = tr.doc.content.size
  const { from, to } = subtreeRange(headings, index, docSize)
  if (from === 0 && to === docSize) {
    const para = schema.nodes.docParagraph?.create()
    if (!para) return false
    tr.replaceWith(0, docSize, para)
  } else tr.delete(from, to)
  tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(from, tr.doc.content.size)), 1))
  return true
}

export function selectHeadingSubtree(tr: Transaction, headings: HeadingRef[], index: number): void {
  const { from, to } = subtreeRange(headings, index, tr.doc.content.size)
  const $from = tr.doc.resolve(from)
  const $to = tr.doc.resolve(to)
  tr.setSelection(TextSelection.between($from, $to))
}
