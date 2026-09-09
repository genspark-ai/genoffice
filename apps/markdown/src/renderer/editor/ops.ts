import type { Editor, ChainedCommands, SingleCommands } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { NodeSelection, TextSelection, type Transaction } from '@tiptap/pm/state'
import type { Mapping } from '@tiptap/pm/transform'
import { createTable } from '@tiptap/extension-table'
import { stripLegacyFencedDivs } from '../markdown/docText'
import type { StringKey } from '../i18n/locale'

/**
 * The single edit entry shared by the AI (`apply_ops`) and the discrete UI
 * actions (ribbon, slash menu, table menu, block handle). Typing and paste
 * stay on the raw ProseMirror step layer; everything with a name goes
 * through here so both callers get the same validation, addressing and
 * result reporting.
 */

export type BlockTarget = 'selection' | { start: number; end?: number }
/** block index to insert after (-1 = document start) or 'selection' = after the caret's block */
export type InsertAnchor = 'selection' | number

export const STYLABLE_MARKS = ['bold', 'italic', 'strike', 'code'] as const
export const BLOCK_TYPES = ['paragraph', 'heading', 'blockquote', 'codeBlock'] as const
export const LIST_KINDS = ['bullet', 'ordered', 'task'] as const
export const TABLE_ACTIONS = [
  'addRowBefore',
  'addRowAfter',
  'deleteRow',
  'addColumnBefore',
  'addColumnAfter',
  'deleteColumn',
  'toggleHeaderRow',
  'deleteTable',
] as const

export type StylableMark = (typeof STYLABLE_MARKS)[number]
export type BlockType = (typeof BLOCK_TYPES)[number]
export type ListKind = (typeof LIST_KINDS)[number]
export type TableAction = (typeof TABLE_ACTIONS)[number]

export type MdOp =
  | { op: 'insertContent'; after: InsertAnchor; markdown: string }
  | { op: 'replaceBlocks'; target: BlockTarget; markdown: string }
  | { op: 'deleteBlocks'; target: BlockTarget }
  | { op: 'replaceText'; target: BlockTarget; find: string; replace: string }
  | {
      op: 'setStyle'
      target: BlockTarget
      style: StylableMark
      find?: string
      mode?: 'apply' | 'remove' | 'toggle'
    }
  | { op: 'setLink'; target: BlockTarget; href: string | null; find?: string }
  | { op: 'setBlockType'; target: BlockTarget; type: BlockType; level?: number; language?: string }
  | { op: 'toggleList'; target: BlockTarget; list: ListKind }
  | { op: 'moveBlocks'; target: BlockTarget; after: number }
  | { op: 'duplicateBlocks'; target: BlockTarget }
  | { op: 'insertTable'; after: InsertAnchor; rows?: number; cols?: number; headerRow?: boolean }
  | { op: 'insertHorizontalRule'; after: InsertAnchor }
  | { op: 'insertImage'; after: InsertAnchor; src: string; alt?: string }
  | { op: 'editTable'; target: BlockTarget; action: TableAction; row?: number; col?: number }
  | { op: 'setFrontmatter'; yaml: string }

export type OpName = MdOp['op']

/** raw-text access to the YAML properties block (inner text, no --- fences) */
export interface FrontmatterAccess {
  read(): string
  /** replace the whole inner YAML; empty string removes the block */
  write(inner: string): void
}

export interface OpsContext {
  source: 'ai' | 'ui'
  frontmatter?: FrontmatterAccess
}

interface Range {
  from: number
  to: number
}

export type OpResult = { ok: true; message: string } | { ok: false; error: string }

/** transaction meta every runOps dispatch carries; plugins (AI highlight, logs) key off it */
export const OP_META = 'mdOp'
export interface OpMeta {
  op: OpName
  source: OpsContext['source']
  /** one id per runOps call */
  batch: number
  /** undoes a failed op's partial changes; not an edit of its own */
  rollback?: boolean
}

export interface RunOpsResult {
  results: OpResult[]
  /** ops executed before the first failure (all of them on success) */
  applied: number
  /** the top-level block list changed shape — indexes the model holds are stale */
  blocksChanged: boolean
}

type FieldType =
  | 'target'
  | 'anchor'
  | 'blockIndex'
  | 'string'
  | 'nullableString'
  | 'int'
  | 'bool'
  | { enum: readonly string[] }

interface FieldSpec {
  type: FieldType
  required?: boolean
  doc: string
}

interface OpSpec {
  labelKey: StringKey
  doc: string
  fields: Record<string, FieldSpec>
}

const TARGET: FieldSpec = { type: 'target', required: true, doc: 'blocks to act on' }
const AFTER: FieldSpec = { type: 'anchor', required: true, doc: 'where to insert' }

export const OP_SPECS: Record<OpName, OpSpec> = {
  insertContent: {
    labelKey: 'aiToolInsert',
    doc: 'insert markdown blocks. On a blank document this replaces the empty paragraph.',
    fields: { after: AFTER, markdown: { type: 'string', required: true, doc: 'GFM content' } },
  },
  replaceBlocks: {
    labelKey: 'aiToolReplace',
    doc: 'replace the target blocks with new markdown (full rewrites, structural changes).',
    fields: {
      target: TARGET,
      markdown: { type: 'string', required: true, doc: 'replacement GFM; empty deletes' },
    },
  },
  deleteBlocks: {
    labelKey: 'blockDelete',
    doc: 'delete the target blocks (the document always keeps at least one paragraph).',
    fields: { target: TARGET },
  },
  replaceText: {
    labelKey: 'aiToolReplaceText',
    doc: 'replace every exact plain-text occurrence inside the target blocks, keeping structure and formatting. Prefer this for small in-place fixes. Matches never cross paragraph or table-cell boundaries.',
    fields: {
      target: TARGET,
      find: { type: 'string', required: true, doc: 'exact plain text, case-sensitive' },
      replace: { type: 'string', required: true, doc: 'plain text; empty deletes' },
    },
  },
  setStyle: {
    labelKey: 'aiToolStyleText',
    doc: 'apply/remove an inline style. With `find`: on every occurrence of that text inside the target; without: on the whole target text.',
    fields: {
      target: TARGET,
      style: { type: { enum: STYLABLE_MARKS }, required: true, doc: 'inline style' },
      find: { type: 'string', doc: 'exact plain text to match' },
      mode: { type: { enum: ['apply', 'remove', 'toggle'] }, doc: 'default apply' },
    },
  },
  setLink: {
    labelKey: 'aiOpSetLink',
    doc: 'link text. With `find`: every occurrence inside the target; without: the whole target text. href null removes the link.',
    fields: {
      target: TARGET,
      href: { type: 'nullableString', required: true, doc: 'URL or null' },
      find: { type: 'string', doc: 'exact plain text to match' },
    },
  },
  setBlockType: {
    labelKey: 'aiOpSetBlockType',
    doc: 'convert the target blocks to paragraph / heading / blockquote / codeBlock (list items are lifted out of their list first).',
    fields: {
      target: TARGET,
      type: { type: { enum: BLOCK_TYPES }, required: true, doc: 'block type' },
      level: { type: 'int', doc: '1-6, heading only' },
      language: { type: 'string', doc: 'codeBlock only' },
    },
  },
  toggleList: {
    labelKey: 'aiOpToggleList',
    doc: 'turn the target blocks into a bullet / ordered / task list, or back into paragraphs when they already are that list.',
    fields: {
      target: TARGET,
      list: { type: { enum: LIST_KINDS }, required: true, doc: 'list kind' },
    },
  },
  moveBlocks: {
    labelKey: 'aiOpMoveBlocks',
    doc: 'move the target blocks to sit after block `after` (-1 = document start).',
    fields: {
      target: TARGET,
      after: { type: 'blockIndex', required: true, doc: 'destination block index' },
    },
  },
  duplicateBlocks: {
    labelKey: 'blockDuplicate',
    doc: 'insert a copy of the target blocks right after them.',
    fields: { target: TARGET },
  },
  insertTable: {
    labelKey: 'insertTable',
    doc: 'insert an empty table (fill cells afterwards with replaceText, or insert a whole pipe table with insertContent instead).',
    fields: {
      after: AFTER,
      rows: { type: 'int', doc: 'default 3' },
      cols: { type: 'int', doc: 'default 3' },
      headerRow: { type: 'bool', doc: 'default true' },
    },
  },
  insertHorizontalRule: {
    labelKey: 'insertHr',
    doc: 'insert a horizontal rule.',
    fields: { after: AFTER },
  },
  insertImage: {
    labelKey: 'aiToolInsertImage',
    doc: 'insert an image block from a path already stored beside the document (insert_image / generate_image do the download and call this for you).',
    fields: {
      after: AFTER,
      src: { type: 'string', required: true, doc: 'relative path or URL' },
      alt: { type: 'string', doc: 'alt text' },
    },
  },
  editTable: {
    labelKey: 'aiOpEditTable',
    doc: 'structural table edit at a cell of the target table block. row/col default 0.',
    fields: {
      target: TARGET,
      action: { type: { enum: TABLE_ACTIONS }, required: true, doc: 'what to do' },
      row: { type: 'int', doc: '0-based row of the cell to act at' },
      col: { type: 'int', doc: '0-based column of the cell to act at' },
    },
  },
  setFrontmatter: {
    labelKey: 'aiToolSetFm',
    doc: 'replace the whole YAML properties block (inner YAML only, no --- fences; empty removes it). Read it first and keep the keys you are not changing.',
    fields: { yaml: { type: 'string', required: true, doc: 'inner YAML' } },
  },
}

function fieldTypeName(type: FieldType): string {
  if (typeof type === 'object') return type.enum.map((v) => `"${v}"`).join('|')
  switch (type) {
    case 'target':
      return 'target'
    case 'anchor':
      return 'int|"selection"'
    case 'blockIndex':
      return 'int'
    case 'nullableString':
      return 'string|null'
    default:
      return type
  }
}

/** model-facing catalogue; one line per op, generated from OP_SPECS so it cannot drift */
export function buildOpsGuide(): string {
  const lines = [
    'Each op is a JSON object with an `op` name plus fields. Addressing: `target` is "selection" (the blocks covered by the user selection) or {start, end?} (0-based inclusive top-level block indexes, end defaults to start); `after` is a block index to insert after (-1 = document start) or "selection" (after the block at the caret). All indexes in one call refer to the document as it was before the call, so ops may be listed in any order.',
    '',
    ...(Object.keys(OP_SPECS) as OpName[]).map((name) => {
      const spec = OP_SPECS[name]
      const fields = Object.entries(spec.fields)
        .map(([f, s]) => `${f}${s.required ? '' : '?'}: ${fieldTypeName(s.type)}`)
        .join(', ')
      return `- ${name} {${fields}} — ${spec.doc}`
    }),
  ]
  return lines.join('\n')
}

function isInt(v: unknown): v is number {
  return typeof v === 'number' && Number.isInteger(v)
}

function validateField(type: FieldType, v: unknown): string | null {
  switch (type) {
    case 'target':
      if (v === 'selection') return null
      if (typeof v !== 'object' || v === null) return 'must be "selection" or {start, end?}'
      {
        const { start, end, ...rest } = v as Record<string, unknown>
        if (!isInt(start) || start < 0) return 'target.start must be a non-negative integer'
        if (end !== undefined && (!isInt(end) || end < start))
          return 'target.end must be an integer >= start'
        if (Object.keys(rest).length)
          return `unknown target field(s): ${Object.keys(rest).join(', ')}`
      }
      return null
    case 'anchor':
      return v === 'selection' || (isInt(v) && v >= -1)
        ? null
        : 'must be an integer >= -1 or "selection"'
    case 'blockIndex':
      return isInt(v) && v >= -1 ? null : 'must be an integer >= -1'
    case 'string':
      return typeof v === 'string' ? null : 'must be a string'
    case 'nullableString':
      return typeof v === 'string' || v === null ? null : 'must be a string or null'
    case 'int':
      return isInt(v) ? null : 'must be an integer'
    case 'bool':
      return typeof v === 'boolean' ? null : 'must be a boolean'
    default:
      return typeof v === 'string' && type.enum.includes(v)
        ? null
        : `must be one of ${type.enum.join(', ')}`
  }
}

export function validateOps(raw: unknown): { ops: MdOp[] } | { error: string } {
  if (!Array.isArray(raw) || raw.length === 0) return { error: 'ops must be a non-empty array' }
  const ops: MdOp[] = []
  for (let i = 0; i < raw.length; i++) {
    const item = raw[i] as Record<string, unknown> | null
    if (typeof item !== 'object' || item === null) return { error: `ops[${i}] must be an object` }
    const name = item.op as OpName
    const spec = OP_SPECS[name]
    if (typeof name !== 'string' || !spec)
      return { error: `ops[${i}]: unknown op "${String(item.op)}"` }
    for (const [field, fs] of Object.entries(spec.fields)) {
      const v = item[field]
      if (v === undefined || v === null) {
        if (fs.required && !(fs.type === 'nullableString' && v === null))
          return { error: `ops[${i}] ${name}: missing required field "${field}"` }
        continue
      }
      const err = validateField(fs.type, v)
      if (err) return { error: `ops[${i}] ${name}: ${field} ${err}` }
    }
    for (const field of Object.keys(item)) {
      if (field !== 'op' && !(field in spec.fields))
        return { error: `ops[${i}] ${name}: unknown field "${field}"` }
    }
    if (name === 'setBlockType') {
      if (item.type === 'heading' && !(isInt(item.level) && item.level >= 1 && item.level <= 6))
        return { error: `ops[${i}] setBlockType: heading needs level 1-6` }
    }
    ops.push(item as unknown as MdOp)
  }
  return { ops }
}

/** true when any op addresses blocks by index (the AI staleness guard applies) */
export function usesBlockIndexes(ops: MdOp[]): boolean {
  return ops.some(
    (op) =>
      ('target' in op && op.target !== 'selection') ||
      ('after' in op && typeof op.after === 'number'),
  )
}

/** top-level block index range covered by [from, to] */
export function blockIndexRange(
  doc: PmNode,
  from: number,
  to: number,
): { startIndex: number; endIndex: number } {
  let startIndex = -1
  let endIndex = -1
  let index = 0
  doc.forEach((node, offset) => {
    if (offset + node.nodeSize > from && offset < to) {
      if (startIndex === -1) startIndex = index
      endIndex = index
    }
    index++
  })
  if (startIndex === -1) {
    startIndex = doc.childCount - 1
    endIndex = startIndex
  }
  return { startIndex, endIndex }
}

/** doc positions [from, to) spanning top-level blocks from..to inclusive */
export function blockRange(doc: PmNode, from: number, to: number): Range {
  let pos = 0
  let start = 0
  let end = 0
  for (let i = 0; i <= to; i++) {
    const child = doc.child(i)
    if (i === from) start = pos
    pos += child.nodeSize
    if (i === to) end = pos
  }
  return { from: start, to: end }
}

/** exactly one empty paragraph — a lone image/math/table with no text is content, not blank */
export function isBlankDoc(doc: PmNode): boolean {
  if (doc.childCount !== 1) return doc.childCount === 0
  return isEmptyParagraph(doc.firstChild!)
}

function isEmptyParagraph(node: PmNode): boolean {
  if (node.type.name !== 'paragraph') return false
  let hasLeaf = false
  node.descendants((n) => {
    if (!n.isText) hasLeaf = true
    return !hasLeaf
  })
  return !hasLeaf && node.textContent.trim() === ''
}

/** [start, end) of the top-level blocks covered by the selection */
export function selectionBlockRange(editor: Editor): Range {
  const sel = editor.state.selection
  if (sel instanceof NodeSelection && sel.$from.depth === 0) return { from: sel.from, to: sel.to }
  const { startIndex, endIndex } = blockIndexRange(editor.state.doc, sel.from, sel.to)
  return blockRange(editor.state.doc, startIndex, endIndex)
}

export function parseMarkdownToNodes(editor: Editor, markdown: string): PmNode[] {
  // model output guard: `:::` fenced divs are not GFM and would land as
  // literal text — strip the fences and keep the body (same as file open).
  // Raw HTML needs no guard: parse runs it through the schema, so semantic
  // tags degrade to their GFM equivalents (<b>→bold, <img>→image) and
  // anything the schema cannot represent loses its styling, keeping text.
  const json = editor.markdown?.parse(stripLegacyFencedDivs(markdown))
  const content = json?.content ?? []
  return content.map((c) => editor.schema.nodeFromJSON(c))
}

/** flattened text of [from, to) with a PM position per character; -1 marks
 *  atom placeholders and block seams so a match can never cross them */
function textIndexOf(doc: PmNode, from: number, to: number): { text: string; pos: number[] } {
  let text = ''
  const pos: number[] = []
  doc.nodesBetween(from, to, (node, p) => {
    if (node.isText && node.text) {
      const start = Math.max(from, p)
      const end = Math.min(to, p + node.text.length)
      for (let i = start; i < end; i++) {
        text += node.text[i - p]
        pos.push(i)
      }
    } else if (node.isLeaf) {
      text += ' '
      pos.push(-1)
    } else if (node.isBlock && text.length > 0 && !text.endsWith('\n')) {
      text += '\n'
      pos.push(-1)
    }
    return true
  })
  return { text, pos }
}

const MAX_MATCHES = 200

/** non-overlapping PM ranges of every plain-text occurrence of `find` in the range */
function findMatches(doc: PmNode, range: Range, find: string): Range[] {
  const index = textIndexOf(doc, range.from, range.to)
  const ranges: Range[] = []
  let at = index.text.indexOf(find)
  while (at !== -1 && ranges.length < MAX_MATCHES) {
    const first = index.pos[at]!
    const last = index.pos[at + find.length - 1]!
    let contiguous = first !== -1 && last - first === find.length - 1
    for (let i = at; contiguous && i < at + find.length; i++) {
      if (index.pos[i] === -1) contiguous = false
    }
    if (contiguous) ranges.push({ from: first, to: last + 1 })
    at = index.text.indexOf(find, at + find.length)
  }
  return ranges
}

class OpError extends Error {}

function fail(msg: string): never {
  throw new OpError(msg)
}

interface Exec {
  editor: Editor
  ctx: OpsContext
  /** index targets/anchors resolve against the doc as it was when the batch started */
  origDoc: PmNode
  mappings: Mapping[]
  meta: OpMeta
}

function mapPos(exec: Exec, pos: number, assoc: -1 | 1): number {
  return exec.mappings.reduce((p, m) => m.map(p, assoc), pos)
}

function describeBlocks(doc: PmNode, range: Range): string {
  const { startIndex, endIndex } = blockIndexRange(doc, range.from, range.to)
  return startIndex === endIndex ? `block ${startIndex}` : `blocks ${startIndex}-${endIndex}`
}

function resolveTarget(exec: Exec, target: BlockTarget): Range {
  const doc = exec.editor.state.doc
  if (target === 'selection') return selectionBlockRange(exec.editor)
  const end = target.end ?? target.start
  if (end > exec.origDoc.childCount - 1)
    fail(`target out of range; the document has ${exec.origDoc.childCount} blocks`)
  const orig = blockRange(exec.origDoc, target.start, end)
  const from = mapPos(exec, orig.from, 1)
  const to = mapPos(exec, orig.to, -1)
  if (to <= from) fail('the target blocks were removed by an earlier op in this batch')
  const { startIndex, endIndex } = blockIndexRange(doc, from, to)
  return blockRange(doc, startIndex, endIndex)
}

/** insertion point plus the empty paragraph to swallow, if any */
function resolveAnchor(exec: Exec, after: InsertAnchor): { pos: number; replace?: Range } {
  const doc = exec.editor.state.doc
  if (isBlankDoc(doc)) return { pos: 0, replace: { from: 0, to: doc.content.size } }
  if (after === 'selection') {
    // a caret inside a list item inserts after that item, not after the whole list
    liftFromList(exec)
    const range = selectionBlockRange(exec.editor)
    const block = exec.editor.state.doc.nodeAt(range.from)
    const lone = block && range.to === range.from + block.nodeSize
    return lone && isEmptyParagraph(block) ? { pos: range.from, replace: range } : { pos: range.to }
  }
  if (after > exec.origDoc.childCount - 1)
    fail(`after out of range; the document has ${exec.origDoc.childCount} blocks`)
  const origPos = after === -1 ? 0 : blockRange(exec.origDoc, after, after).to
  return { pos: mapPos(exec, origPos, 1) }
}

/** invert a dispatched transaction step by step (highlights inside it fall away with the content) */
function rollback(exec: Exec, applied: Transaction): void {
  const tr = exec.editor.state.tr
  for (let i = applied.steps.length - 1; i >= 0; i--) {
    tr.step(applied.steps[i]!.invert(applied.docs[i]!))
  }
  tr.setMeta(OP_META, { ...exec.meta, rollback: true })
  exec.editor.view.dispatch(tr)
}

function nodesSize(nodes: PmNode[]): number {
  return nodes.reduce((s, n) => s + n.nodeSize, 0)
}

function dispatch(exec: Exec, tr: Transaction): void {
  tr.setMeta(OP_META, exec.meta)
  exec.editor.view.dispatch(exec.ctx.source === 'ui' ? tr.scrollIntoView() : tr)
}

function insertNodes(exec: Exec, after: InsertAnchor, nodes: PmNode[]): void {
  const { pos, replace } = resolveAnchor(exec, after)
  let tr = exec.editor.state.tr
  tr = replace ? tr.replaceWith(replace.from, replace.to, nodes) : tr.insert(pos, nodes)
  if (exec.ctx.source === 'ui') {
    const end = pos + nodesSize(nodes)
    // a trailing atom (rule/image) leaves nowhere to type — give the caret a paragraph
    if (!nodes.at(-1)!.isTextblock && end >= tr.doc.content.size) {
      tr = tr.insert(end, exec.editor.schema.nodes.paragraph!.create())
    }
    tr = tr.setSelection(
      TextSelection.near(tr.doc.resolve(Math.min(pos + 1, tr.doc.content.size)), 1),
    )
  }
  dispatch(exec, tr)
}

/** put the selection on the range (index targets) so tiptap commands act on it; no-op for 'selection' */
function requireText(doc: PmNode, range: Range): void {
  let hasText = false
  doc.nodesBetween(range.from, range.to, (node) => {
    if (node.isTextblock) hasText = true
    return !hasText
  })
  if (!hasText) fail(`${describeBlocks(doc, range)} has no text to act on (image/rule blocks)`)
}

function selectTarget(exec: Exec, target: BlockTarget, range: Range): void {
  if (target === 'selection') return
  const doc = exec.editor.state.doc
  requireText(doc, range)
  const sel = TextSelection.between(doc.resolve(range.from), doc.resolve(range.to))
  dispatch(exec, exec.editor.state.tr.setSelection(sel))
}

function chain(exec: Exec): ChainedCommands {
  const c = exec.editor.chain().command(({ tr }) => {
    tr.setMeta(OP_META, exec.meta)
    return true
  })
  return exec.ctx.source === 'ui' ? c.focus() : c
}

/**
 * Block-type conversions are illegal inside a list item — setHeading would
 * silently fail. Lift the current item out of its list(s) first.
 */
function liftFromList(exec: Exec): void {
  const { editor } = exec
  for (let guard = 0; guard < 10; guard++) {
    const itemName = editor.isActive('taskItem')
      ? 'taskItem'
      : editor.isActive('listItem')
        ? 'listItem'
        : null
    if (!itemName) return
    if (!chain(exec).liftListItem(itemName).run()) return
  }
}

const LIST_NODES = new Set(['bulletList', 'orderedList', 'taskList'])

/**
 * Replace every top-level list (and blockquote) in the range with the blocks
 * inside it so a block-type conversion can reach them — the index-addressed
 * twin of liftFromList, which only lifts the item at the caret.
 */
function flattenWrappers(tr: Transaction, range: Range): Range {
  for (;;) {
    const doc = tr.doc
    const wrappers: Array<{ from: number; node: PmNode }> = []
    doc.nodesBetween(range.from, range.to, (node, pos, parent) => {
      if (parent === doc && (LIST_NODES.has(node.type.name) || node.type.name === 'blockquote'))
        wrappers.push({ from: pos, node })
      return false
    })
    if (wrappers.length === 0) return range
    for (const { from, node } of wrappers.reverse()) {
      const blocks: PmNode[] = []
      if (node.type.name === 'blockquote') node.forEach((child) => blocks.push(child))
      else node.forEach((item) => item.forEach((child) => blocks.push(child)))
      tr.replaceWith(from, from + node.nodeSize, blocks)
    }
    range = { from: range.from, to: range.to + tr.doc.content.size - doc.content.size }
  }
}

/** every top-level block in the (flattened) range already is the requested type — tiptap reports a no-op as failure */
function alreadyTyped(
  doc: PmNode,
  range: Range,
  op: Extract<MdOp, { op: 'setBlockType' }>,
): boolean {
  if (op.type === 'blockquote') return false
  let all = true
  doc.nodesBetween(range.from, range.to, (node, _pos, parent) => {
    if (parent !== doc) return false
    if (node.type.name !== op.type) all = false
    else if (op.type === 'heading' && node.attrs.level !== op.level) all = false
    else if (op.type === 'codeBlock' && op.language && node.attrs.language !== op.language)
      all = false
    return false
  })
  return all
}

/** caret-scoped unwrap of blockquotes (setBlockType paragraph = plain paragraph) */
function liftFromQuote(exec: Exec): void {
  for (let guard = 0; guard < 10 && exec.editor.isActive('blockquote'); guard++) {
    if (!chain(exec).lift('blockquote').run()) return
  }
}

/** run a selection-based tiptap command on the target; returns the target's label for the result message */
function withSelection(exec: Exec, target: BlockTarget, run: () => boolean): string {
  const range = resolveTarget(exec, target)
  const label = describeBlocks(exec.editor.state.doc, range)
  selectTarget(exec, target, range)
  if (!run()) fail('the editor rejected this edit at the target')
  return label
}

function markRanges(
  exec: Exec,
  ranges: Range[],
  markName: string,
  attrs: Record<string, unknown> | null,
  mode: 'apply' | 'remove' | 'toggle',
): 'applied' | 'removed' {
  const markType = exec.editor.schema.marks[markName]!
  const doc = exec.editor.state.doc
  const allMarked = ranges.every((r) => doc.rangeHasMark(r.from, r.to, markType))
  const remove = mode === 'remove' || (mode === 'toggle' && allMarked)
  let tr = exec.editor.state.tr
  for (const r of ranges) {
    tr = remove
      ? tr.removeMark(r.from, r.to, markType)
      : tr.addMark(r.from, r.to, markType.create(attrs))
  }
  dispatch(exec, tr)
  return remove ? 'removed' : 'applied'
}

function matchesIn(exec: Exec, range: Range, find: string): Range[] {
  if (!find) fail('find must not be empty')
  const doc = exec.editor.state.doc
  const matches = findMatches(doc, range, find)
  if (matches.length === 0) {
    fail(
      `"${find}" was not found in ${describeBlocks(doc, range)}. The match is exact plain text (check spacing and punctuation); read the blocks to see their current text.`,
    )
  }
  return matches
}

function cellTextPos(table: PmNode, tablePos: number, row: number, col: number): number {
  if (row < 0 || row >= table.childCount)
    fail(`row out of range; the table has ${table.childCount} rows`)
  const rowNode = table.child(row)
  if (col < 0 || col >= rowNode.childCount)
    fail(`col out of range; the row has ${rowNode.childCount} cells`)
  let pos = tablePos + 1
  for (let r = 0; r < row; r++) pos += table.child(r).nodeSize
  pos += 1
  for (let c = 0; c < col; c++) pos += rowNode.child(c).nodeSize
  return pos + 2
}

function execOp(exec: Exec, op: MdOp): OpResult {
  const { editor } = exec
  switch (op.op) {
    case 'insertContent': {
      if (!op.markdown.trim()) fail('markdown must not be empty')
      const nodes = parseMarkdownToNodes(editor, op.markdown)
      if (nodes.length === 0) fail('markdown parsed to no content')
      insertNodes(exec, op.after, nodes)
      return { ok: true, message: `Inserted ${nodes.length} block(s).` }
    }

    case 'replaceBlocks':
    case 'deleteBlocks': {
      const nodes = op.op === 'replaceBlocks' ? parseMarkdownToNodes(editor, op.markdown) : []
      const range = resolveTarget(exec, op.target)
      const doc = editor.state.doc
      const label = describeBlocks(doc, range)
      let tr = editor.state.tr
      if (nodes.length === 0) {
        // the schema needs one block — deleting everything leaves an empty paragraph
        tr =
          range.from === 0 && range.to === doc.content.size
            ? tr.replaceWith(range.from, range.to, editor.schema.nodes.paragraph!.create())
            : tr.delete(range.from, range.to)
        if (exec.ctx.source === 'ui') {
          tr = tr.setSelection(
            TextSelection.near(tr.doc.resolve(Math.min(range.from, tr.doc.content.size))),
          )
        }
        dispatch(exec, tr)
        return { ok: true, message: `Deleted ${label}.` }
      }
      tr = tr.replaceWith(range.from, range.to, nodes)
      dispatch(exec, tr)
      return { ok: true, message: `Replaced ${label} with ${nodes.length} block(s).` }
    }

    case 'replaceText': {
      const range = resolveTarget(exec, op.target)
      const label = describeBlocks(editor.state.doc, range)
      const matches = matchesIn(exec, range, op.find)
      let tr = editor.state.tr
      // bottom-up so earlier matches keep their positions while editing
      for (let i = matches.length - 1; i >= 0; i--) {
        const m = matches[i]!
        tr = op.replace ? tr.insertText(op.replace, m.from, m.to) : tr.delete(m.from, m.to)
      }
      dispatch(exec, tr)
      return { ok: true, message: `Replaced ${matches.length} occurrence(s) in ${label}.` }
    }

    case 'setStyle': {
      const mode = op.mode ?? 'apply'
      if (op.find !== undefined) {
        const range = resolveTarget(exec, op.target)
        const matches = matchesIn(exec, range, op.find)
        const done = markRanges(exec, matches, op.style, null, mode)
        return {
          ok: true,
          message: `${done === 'applied' ? 'Applied' : 'Removed'} ${op.style} on ${matches.length} match(es).`,
        }
      }
      const label = withSelection(exec, op.target, () => {
        const c = chain(exec)
        const cmd =
          mode === 'toggle'
            ? c.toggleMark(op.style)
            : mode === 'remove'
              ? c.unsetMark(op.style)
              : c.setMark(op.style)
        return cmd.run()
      })
      return { ok: true, message: `Set ${op.style} (${mode}) on ${label}.` }
    }

    case 'setLink': {
      const attrs = op.href === null ? null : { href: op.href }
      if (op.find !== undefined) {
        const range = resolveTarget(exec, op.target)
        const matches = matchesIn(exec, range, op.find)
        markRanges(exec, matches, 'link', attrs, attrs ? 'apply' : 'remove')
        return {
          ok: true,
          message: `${attrs ? 'Linked' : 'Unlinked'} ${matches.length} match(es).`,
        }
      }
      const label = withSelection(exec, op.target, () => {
        const c = chain(exec).extendMarkRange('link')
        return (attrs ? c.setLink(attrs) : c.unsetLink()).run()
      })
      return { ok: true, message: `${attrs ? 'Linked' : 'Unlinked'} the text of ${label}.` }
    }

    case 'setBlockType': {
      const name = op.type === 'heading' ? `h${op.level}` : op.type
      const convert = (commands: SingleCommands): boolean => {
        switch (op.type) {
          case 'paragraph':
            return commands.setParagraph()
          case 'heading':
            return commands.setHeading({ level: op.level as 1 | 2 | 3 | 4 | 5 | 6 })
          case 'blockquote':
            return commands.setBlockquote()
          case 'codeBlock':
            return commands.setCodeBlock(op.language ? { language: op.language } : undefined)
        }
      }
      if (op.target === 'selection') {
        const label = withSelection(exec, op.target, () => {
          // already quoted = nothing to do; lifting one paragraph would split a multi-paragraph quote
          if (op.type === 'blockquote' && editor.isActive('blockquote')) return true
          liftFromList(exec)
          if (op.type === 'paragraph') liftFromQuote(exec)
          const c = chain(exec)
          return (op.type === 'blockquote' ? c.setParagraph() : c)
            .command(({ commands }) => convert(commands))
            .run()
        })
        return { ok: true, message: `Set ${label} to ${name}.` }
      }
      // index target: unwrap lists/quotes in range and convert in one transaction, so a
      // failure never leaves a half-unwrapped document behind
      const range = resolveTarget(exec, op.target)
      const label = describeBlocks(editor.state.doc, range)
      requireText(editor.state.doc, range)
      let flat = range
      let applied: Transaction | null = null
      const capture = ({ transaction }: { transaction: Transaction }) => {
        if (transaction.docChanged) applied = transaction
      }
      editor.on('transaction', capture)
      // two chain steps: command props snapshot the selection when the step starts, so the
      // conversion must run in a step after the one that moved the selection
      const ok = chain(exec)
        .command(({ tr }) => {
          flat = flattenWrappers(tr, range)
          tr.setSelection(TextSelection.between(tr.doc.resolve(flat.from), tr.doc.resolve(flat.to)))
          return true
        })
        .command(({ tr, commands }) => alreadyTyped(tr.doc, flat, op) || convert(commands))
        .run()
      editor.off('transaction', capture)
      if (!ok) {
        if (applied) rollback(exec, applied)
        fail('the editor rejected this conversion at the target')
      }
      return { ok: true, message: `Set ${label} to ${name}.` }
    }

    case 'toggleList': {
      const label = withSelection(exec, op.target, () => {
        const c = chain(exec)
        switch (op.list) {
          case 'bullet':
            return c.toggleBulletList().run()
          case 'ordered':
            return c.toggleOrderedList().run()
          case 'task':
            return c.toggleTaskList().run()
        }
      })
      return { ok: true, message: `Toggled ${op.list} list on ${label}.` }
    }

    case 'moveBlocks': {
      const range = resolveTarget(exec, op.target)
      const { pos: anchor } = resolveAnchor(exec, op.after)
      if (anchor > range.from && anchor < range.to)
        fail('after must not point inside the moved blocks')
      if (anchor === range.from || anchor === range.to)
        return { ok: true, message: 'The blocks are already there.' }
      const content = editor.state.doc.slice(range.from, range.to).content
      const caretOffset = editor.state.selection.head - range.from
      const insertAt = anchor > range.to ? anchor - content.size : anchor
      let tr = editor.state.tr.delete(range.from, range.to).insert(insertAt, content)
      if (exec.ctx.source === 'ui' && caretOffset >= 0 && caretOffset <= content.size) {
        tr = tr.setSelection(TextSelection.near(tr.doc.resolve(insertAt + caretOffset)))
      }
      dispatch(exec, tr)
      return {
        ok: true,
        message: `Moved the blocks; they are now ${describeBlocks(tr.doc, { from: insertAt, to: insertAt + content.size })}.`,
      }
    }

    case 'duplicateBlocks': {
      const range = resolveTarget(exec, op.target)
      const label = describeBlocks(editor.state.doc, range)
      const content = editor.state.doc.slice(range.from, range.to).content
      let tr = editor.state.tr.insert(range.to, content)
      if (exec.ctx.source === 'ui') {
        const caretOffset = editor.state.selection.head - range.from
        tr = tr.setSelection(
          TextSelection.near(tr.doc.resolve(range.to + Math.max(0, caretOffset))),
        )
      }
      dispatch(exec, tr)
      return { ok: true, message: `Duplicated ${label}.` }
    }

    case 'insertTable': {
      const rows = op.rows ?? 3
      const cols = op.cols ?? 3
      if (rows < 1 || cols < 1 || rows > 100 || cols > 30) fail('rows must be 1-100 and cols 1-30')
      const table = createTable(editor.schema, rows, cols, op.headerRow ?? true)
      insertNodes(exec, op.after, [table])
      return { ok: true, message: `Inserted a ${rows}x${cols} table.` }
    }

    case 'insertHorizontalRule': {
      insertNodes(exec, op.after, [editor.schema.nodes.horizontalRule!.create()])
      return { ok: true, message: 'Inserted a horizontal rule.' }
    }

    case 'insertImage': {
      if (!op.src.trim()) fail('src must not be empty')
      const node = editor.schema.nodes.image!.create({ src: op.src, alt: op.alt?.trim() || null })
      insertNodes(exec, op.after, [node])
      return { ok: true, message: 'Inserted the image.' }
    }

    case 'editTable': {
      const range = resolveTarget(exec, op.target)
      if (op.target === 'selection') {
        if (!editor.isActive('table')) fail('the selection is not inside a table')
      } else {
        const doc = editor.state.doc
        const table = doc.nodeAt(range.from)
        if (!table || table.type.name !== 'table' || range.to !== range.from + table.nodeSize)
          fail(`${describeBlocks(doc, range)} is not a single table block`)
        const pos = cellTextPos(table, range.from, op.row ?? 0, op.col ?? 0)
        dispatch(exec, editor.state.tr.setSelection(TextSelection.near(doc.resolve(pos))))
      }
      const c = chain(exec)
      const commands: Record<TableAction, () => ChainedCommands> = {
        addRowBefore: () => c.addRowBefore(),
        addRowAfter: () => c.addRowAfter(),
        deleteRow: () => c.deleteRow(),
        addColumnBefore: () => c.addColumnBefore(),
        addColumnAfter: () => c.addColumnAfter(),
        deleteColumn: () => c.deleteColumn(),
        toggleHeaderRow: () => c.toggleHeaderRow(),
        deleteTable: () => c.deleteTable(),
      }
      if (!commands[op.action]().run()) fail(`the table rejected ${op.action}`)
      return { ok: true, message: `Table: ${op.action} done.` }
    }

    case 'setFrontmatter': {
      if (!exec.ctx.frontmatter) fail('frontmatter is not available')
      const inner = op.yaml
      exec.ctx.frontmatter.write(inner.trim() ? inner.replace(/\n+$/, '') : '')
      return { ok: true, message: inner.trim() ? 'Frontmatter updated.' : 'Frontmatter removed.' }
    }
  }
}

/** ops that never add, remove or reorder top-level blocks */
const IN_PLACE_OPS = new Set<OpName>(['replaceText', 'setStyle', 'setLink', 'setFrontmatter'])

/**
 * Execute ops in order, stopping at the first failure. Index-addressed ops
 * keep meaning "the document as it was before the call": positions are
 * mapped through every transaction the earlier ops dispatched.
 */
let batchSeq = 0

export function runOps(editor: Editor, ops: MdOp[], ctx: OpsContext): RunOpsResult {
  const batch = ++batchSeq
  const exec: Exec = {
    editor,
    ctx,
    origDoc: editor.state.doc,
    mappings: [],
    meta: { op: ops[0]!.op, source: ctx.source, batch },
  }
  const onTransaction = ({ transaction }: { transaction: Transaction }) => {
    if (transaction.docChanged) exec.mappings.push(transaction.mapping)
  }
  editor.on('transaction', onTransaction)
  const results: OpResult[] = []
  try {
    for (const op of ops) {
      exec.meta = { op: op.op, source: ctx.source, batch }
      let result: OpResult
      try {
        result = execOp(exec, op)
      } catch (err) {
        if (!(err instanceof OpError)) throw err
        result = { ok: false, error: err.message }
      }
      results.push(result)
      if (!result.ok) break
    }
  } finally {
    editor.off('transaction', onTransaction)
  }
  const applied = results.filter((r) => r.ok).length
  return {
    results,
    applied,
    blocksChanged: ops.slice(0, applied).some((op) => !IN_PLACE_OPS.has(op.op)),
  }
}

/** UI convenience: one op on the current selection, focus kept in the editor */
export function uiOp(editor: Editor, op: MdOp): boolean {
  return runOps(editor, [op], { source: 'ui' }).applied === 1
}

/** swap the selection's top-level blocks with the previous/next sibling; false at the edges */
export function moveSelectedBlocks(editor: Editor, direction: -1 | 1): boolean {
  const { doc, selection } = editor.state
  const { startIndex, endIndex } = blockIndexRange(doc, selection.from, selection.to)
  const after = direction === -1 ? startIndex - 2 : endIndex + 1
  if (after < -1 || after >= doc.childCount) return false
  return uiOp(editor, { op: 'moveBlocks', target: 'selection', after })
}
