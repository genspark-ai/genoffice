/** Word's Table > Convert: paragraphs into a table's cells, and a table back into paragraphs */
import { RANGE_BOOKMARK_MARK } from './bookmark-range'
import type { Editor } from '@tiptap/core'
import type { TableAutoFitMode, TableModel, TextboxDisplay } from '@genoffice/docx-engine'
import { Fragment, type Node as PmNode } from '@tiptap/pm/model'
import { TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { tableModelToPmNode } from './convert'
import { shapeWrapOf } from './floating-z-order'
import { settleCarried } from './object-group'
import type { GroupMemberSource } from './object-multi-select'
import { wordTableCell, wordTableModel } from './new-table'
import { TRACK_IGNORE, activeTracking, revisionDateNow, withRowRecord } from './revisions'
import { setTableAutoFit } from './table-properties'

/** what parts the text into cells, or joins the cells into text: a character, or the paragraph mark */
export type ConvertSeparator = { kind: 'paragraph' } | { kind: 'text'; text: string }

/** an empty Other character parts nothing: each paragraph is a cell */
const byParagraph = (
  separator: ConvertSeparator,
): separator is { kind: 'paragraph' } | { kind: 'text'; text: '' } =>
  separator.kind === 'paragraph' || !separator.text

/** Word's Insert Table column limit */
export const MAX_CONVERT_COLS = 63

interface Block {
  node: PmNode
  pos: number
}

/** the body paragraphs the selection touches; null when it holds anything else (a table, a field block) */
function selectedParagraphs(state: EditorState): Block[] | null {
  const { doc, selection } = state
  const blocks: Block[] = []
  let plain = true
  doc.forEach((node, pos) => {
    const end = pos + node.nodeSize
    const inRange = selection.empty
      ? pos <= selection.from && selection.from < end
      : pos < selection.to && end > selection.from
    if (!inRange) return
    if (!node.isTextblock) plain = false
    blocks.push({ node, pos })
  })
  return plain && blocks.length > 0 ? blocks : null
}

/** a paragraph's inline content cut at each separator; the separators themselves go */
function fieldsOf(node: PmNode, separator: string): Fragment[] {
  const cuts: number[] = []
  node.forEach((child, offset) => {
    if (!child.isText || !child.text) return
    let at = child.text.indexOf(separator)
    while (at >= 0) {
      cuts.push(offset + at)
      at = child.text.indexOf(separator, at + separator.length)
    }
  })
  const fields: Fragment[] = []
  let from = 0
  for (const cut of cuts) {
    fields.push(node.content.cut(from, cut))
    from = cut + separator.length
  }
  fields.push(node.content.cut(from))
  return fields
}

interface Field {
  content: Fragment
  /** the paragraph the field was cut from */
  source: PmNode
  /** the paragraph's first field: what names the paragraph stays on it */
  first: boolean
}

function cellFields(blocks: readonly Block[], separator: ConvertSeparator): Field[][] {
  return blocks.map(({ node }) =>
    (byParagraph(separator) ? [node.content] : fieldsOf(node, separator.text)).map(
      (content, index) => ({ content, source: node, first: index === 0 }),
    ),
  )
}

/** what belongs to the paragraph's place in the body, not to how it is formatted */
const NOT_CARRIED = new Set([
  'docxIndex',
  'tblPara',
  'aiChanged',
  'permEditable',
  'commentStarts',
  'commentEnds',
  'pageBreakBefore',
  'dropCap',
  'frame',
  'frameBox',
  'sdtShell',
  'collapsed',
  'moveRevision',
  'pPrChange',
  'numberingChange',
  'sectPrChange',
  'paraMarkDel',
  'paraMarkIns',
  'paraMarkRPrChange',
  'blockRevision',
  'caretMarks',
  'bookmarksHeld',
])
const NAMES = new Set(['bookmarks', 'hiddenBookmarks'])

/**
 * Struck content hands its bookmarks to what takes its place, so accepting
 * the conversion keeps them; it holds on to the names for a rejection.
 */
function withNamesHeld(node: PmNode): PmNode {
  // the range a bookmark covers goes with the content too: its two ends are written once
  if (node.isInline)
    return node.mark(node.marks.filter((mark) => mark.type.name !== RANGE_BOOKMARK_MARK))
  const children: PmNode[] = []
  node.forEach((child) => children.push(withNamesHeld(child)))
  const content = Fragment.from(children)
  if (!node.isTextblock) return node.copy(content)
  const { bookmarks, hiddenBookmarks } = node.attrs as Record<string, string[] | null | undefined>
  if (!bookmarks?.length && !hiddenBookmarks?.length) return node.copy(content)
  return node.type.create(
    {
      ...node.attrs,
      bookmarks: null,
      hiddenBookmarks: null,
      bookmarksHeld: {
        ...(bookmarks?.length ? { bookmarks } : {}),
        ...(hiddenBookmarks?.length ? { hiddenBookmarks } : {}),
      },
    },
    content,
    node.marks,
  )
}

/**
 * The cell's paragraph for a field: formatted as the paragraph it came from
 * (Word keeps alignment, indents, spacing, style and list), over the new
 * table's own cell paragraph where the source sets nothing.
 */
function cellParagraph(blank: PmNode, field: Field): PmNode {
  const { source } = field
  const type = source.type.name === 'docListItem' ? source.type : blank.type
  const attrs: Record<string, unknown> = {}
  for (const name of Object.keys(type.spec.attrs ?? {})) {
    const own = source.attrs[name]
    const carried =
      name in source.attrs && !NOT_CARRIED.has(name) && (field.first || !NAMES.has(name))
    const set = own !== null && own !== undefined && own !== type.spec.attrs![name].default
    if (carried && set) attrs[name] = own
    else if (name in blank.attrs) attrs[name] = blank.attrs[name]
  }
  return type.create(attrs, field.content)
}

/** the selected paragraphs' texts, for the dialog to count columns from; null when there is nothing to convert */
export function textToTableScope(state: EditorState): string[] | null {
  return selectedParagraphs(state)?.map((block) => block.node.textContent) ?? null
}

/** the column count Word proposes: the most fields any paragraph has; one for paragraph marks */
export function suggestedColumns(texts: readonly string[], separator: ConvertSeparator): number {
  if (byParagraph(separator)) return 1
  const most = Math.max(1, ...texts.map((text) => text.split(separator.text).length))
  return Math.min(MAX_CONVERT_COLS, most)
}

/** how many rows the text makes at that column count */
export function convertedRows(
  texts: readonly string[],
  separator: ConvertSeparator,
  cols: number,
): number {
  if (byParagraph(separator)) return Math.ceil(texts.length / cols)
  return texts.reduce((rows, text) => rows + Math.ceil(text.split(separator.text).length / cols), 0)
}

interface Reviser {
  author: string
  date: string
}

/** a table whose rows are all tracked as inserted or deleted, as Word records a table that comes or goes */
function withRowRecords(table: PmNode, kind: 'ins' | 'del', by: Reviser): PmNode {
  const rows: PmNode[] = []
  table.forEach((row) =>
    rows.push(
      row.attrs.rowRevision
        ? row
        : row.type.create(
            {
              ...row.attrs,
              rowRevision: { kind, ...by },
              rawTrPr: withRowRecord(row.attrs.rawTrPr, kind, by.author, by.date),
            },
            row.content,
          ),
    ),
  )
  return table.type.create(table.attrs, rows)
}

/** the blocks at `pos` come in as a tracked insertion: their text, each paragraph mark, each table row */
function insertTracked(tr: Transaction, pos: number, blocks: PmNode[], by: Reviser): void {
  const { marks } = tr.doc.type.schema
  const nodes = blocks.map((block) =>
    block.type.name === 'docTable'
      ? withRowRecords(block, 'ins', by)
      : block.type.create({ ...block.attrs, paraMarkIns: JSON.stringify(by) }, block.content),
  )
  tr.insert(pos, nodes)
  const size = nodes.reduce((sum, node) => sum + node.nodeSize, 0)
  tr.addMark(pos, pos + size, marks.ins.create(by))
}

/** the text in [from, to) goes as a tracked deletion; what the same reviser inserted there just goes */
function strikeTracked(tr: Transaction, from: number, to: number, by: Reviser): void {
  const { marks } = tr.doc.type.schema
  const own: Array<[number, number]> = []
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true
    const inserted = node.marks.find((mark) => mark.type === marks.ins)
    if (inserted?.attrs.author === by.author) own.push([pos, pos + node.nodeSize])
    return false
  })
  tr.addMark(from, to, marks.del.create(by))
  for (const [start, end] of own.reverse()) tr.delete(start, end)
}

export interface TextToTableOptions {
  cols: number
  separator: ConvertSeparator
  autoFit: TableAutoFitMode
}

/**
 * Turns the selected paragraphs into a table: each paragraph is a row, cut
 * into cells at the separator (with paragraph marks, each paragraph is a
 * cell). Fields past the last column carry on in a new row. False when the
 * selection is not plain paragraphs.
 */
export function convertTextToTable(
  editor: Editor,
  options: TextToTableOptions,
  contentWidthPx: number,
): boolean {
  const { state } = editor
  const blocks = selectedParagraphs(state)
  if (!blocks) return false
  const cols = Math.min(MAX_CONVERT_COLS, Math.max(1, Math.round(options.cols)))
  const track = activeTracking(state)
  const perBlock = cellFields(blocks, options.separator)
  const rows: Field[][] = []
  if (byParagraph(options.separator)) {
    const all = perBlock.flat()
    for (let at = 0; at < all.length; at += cols) rows.push(all.slice(at, at + cols))
  } else {
    for (const fields of perBlock)
      for (let at = 0; at < fields.length; at += cols) rows.push(fields.slice(at, at + cols))
  }
  const model = wordTableModel(
    rows.map(() => Array.from({ length: cols }, () => wordTableCell(['']))),
    cols,
  )
  const empty = state.schema.nodeFromJSON(tableModelToPmNode(model))
  const filled: PmNode[] = []
  empty.forEach((row, _offset, rowIndex) => {
    const cells: PmNode[] = []
    row.forEach((cell, _cellOffset, cellIndex) => {
      const field = rows[rowIndex]?.[cellIndex]
      const blank = cell.firstChild
      cells.push(field && blank ? cell.copy(Fragment.from(cellParagraph(blank, field))) : cell)
    })
    filled.push(row.copy(Fragment.from(cells)))
  })
  const table = empty.copy(Fragment.from(filled))
  const from = blocks[0].pos
  const last = blocks[blocks.length - 1]
  const to = last.pos + last.node.nodeSize
  const tr = state.tr
  if (track) {
    // the text stays, struck, in front of the table that takes its place
    const by = { author: track.author, date: revisionDateNow() }
    insertTracked(tr, to, [table], by)
    for (const { node, pos } of [...blocks].reverse()) {
      // the struck paragraph is still there: a name stands once, on the cell that took it
      tr.setNodeMarkup(pos, undefined, {
        ...withNamesHeld(node).attrs,
        paraMarkDel: JSON.stringify(by),
      })
      const range = state.schema.marks[RANGE_BOOKMARK_MARK]
      if (range) tr.removeMark(pos + 1, pos + node.nodeSize - 1, range)
      strikeTracked(tr, pos + 1, pos + node.nodeSize - 1, by)
    }
    tr.setSelection(TextSelection.near(tr.doc.resolve(tr.mapping.map(to, -1) + 1)))
    tr.setMeta(TRACK_IGNORE, true)
  } else {
    tr.replaceWith(from, to, table)
    tr.setSelection(TextSelection.near(tr.doc.resolve(from + 1)))
  }
  editor.view.dispatch(tr.scrollIntoView())
  if (options.autoFit !== 'window') {
    // rides the same undo step as the conversion
    setTableAutoFit(options.autoFit, contentWidthPx)(editor.state, (fit) =>
      editor.view.dispatch(fit.setMeta('appendedTransaction', tr).setMeta(TRACK_IGNORE, !!track)),
    )
  }
  return true
}

const isCell = (node: PmNode) =>
  node.type.name === 'docTableCell' || node.type.name === 'docTableHeader'

/** the table the selection sits in */
function tableAround(state: EditorState): Block | null {
  const { $from } = state.selection
  for (let depth = $from.depth; depth > 0; depth--) {
    const node = $from.node(depth)
    if (node.type.name === 'docTable') return { node, pos: $from.before(depth) }
  }
  return null
}

export interface TableToTextOptions {
  separator: ConvertSeparator
  /** a table inside a cell becomes text too; otherwise it stays a table among the paragraphs */
  nested: boolean
}

const isCarried = (node: PmNode) => node.type.name === 'docProtected' && !!node.attrs.carried

/**
 * The shapes anchored in a cell paragraph as nodes that follow the paragraph
 * the text of it lands in: carried, so they take no line and are saved as
 * anchors of that paragraph, from the drawings the table held.
 */
function liftedShapes(cellBoxes: PmNode): PmNode[] {
  const boxes = (cellBoxes.attrs.boxes as TextboxDisplay[] | null) ?? []
  const xml = (cellBoxes.attrs.xml as (string | null)[] | null) ?? []
  const type = cellBoxes.type.schema.nodes.docProtected
  const out: PmNode[] = []
  for (const source of new Set(xml)) {
    if (!source || !type) continue
    // on the page now, not in a cell
    const own = boxes
      .filter((_, i) => xml[i] === source)
      .map(({ outsideCell: _outsideCell, ...box }) => box)
    const xEmu = own[0].offsetXEmu ?? 0
    const yEmu = own[0].offsetYEmu ?? 0
    const member: GroupMemberSource = { xml: source, boxes: own, xEmu, yEmu, behindText: true }
    out.push(
      type.create({
        docxIndex: null,
        blockType: 'passthrough',
        label: 'Shape',
        member,
        carried: true,
        textboxes: own,
        imageWrap: shapeWrapOf(own[0]),
        imageZOrder: own[0].z ?? null,
        imageOffsetXEmu: xEmu,
        imageOffsetYEmu: yEmu,
      }),
    )
  }
  return out
}

/** one row's cells in a line: the last paragraph of a cell runs on into the first of the next */
function joinCells(cells: PmNode[][], separator: ConvertSeparator, paragraph: PmNode): PmNode[] {
  if (separator.kind === 'paragraph') return cells.flat()
  const text = separator.text ? paragraph.type.schema.text(separator.text) : null
  const line: PmNode[] = []
  cells.forEach((blocks, index) => {
    const [first, ...rest] = blocks.length > 0 ? blocks : [paragraph]
    // shapes carried behind a paragraph stay behind it as it runs on
    let at = line.length - 1
    while (at >= 0 && isCarried(line[at])) at--
    const prev = line[at]
    if (index === 0 || !prev?.isTextblock || !first.isTextblock) {
      line.push(first, ...rest)
      return
    }
    const joined = text ? prev.content.addToEnd(text) : prev.content
    line[at] = prev.copy(joined.append(first.content))
    line.push(...rest)
  })
  return line
}

function nestedAsText(model: TableModel, separator: ConvertSeparator, paragraph: PmNode): PmNode[] {
  const schema = paragraph.type.schema
  const plain = (text: string) => paragraph.copy(text ? Fragment.from(schema.text(text)) : null)
  return model.rows.flatMap((row) =>
    joinCells(
      row.filter((cell) => cell.vMerge !== 'continue').map((cell) => cell.paras.map(plain)),
      separator,
      paragraph,
    ),
  )
}

/**
 * Replaces the table at the selection with its text: a paragraph per row,
 * the cells joined by the separator (with paragraph marks, a paragraph per
 * cell). False when the selection is in no table.
 */
export function convertTableToText(editor: Editor, options: TableToTextOptions): boolean {
  const { state } = editor
  const target = tableAround(state)
  const paragraphType = state.schema.nodes.docParagraph
  if (!target || !paragraphType) return false
  const blank = paragraphType.create()
  const blocks: PmNode[] = []
  const track = activeTracking(state)
  target.node.forEach((row) => {
    const cells: PmNode[][] = []
    row.forEach((cell) => {
      if (!isCell(cell) || cell.attrs.gridGap || cell.attrs.mergeGap) return
      const content: PmNode[] = []
      // shapes stand ahead of the paragraph they are anchored in
      let shapes: PmNode[] = []
      cell.forEach((child) => {
        if (child.isTextblock) {
          content.push(child, ...shapes)
          shapes = []
        }
        // a tracked conversion keeps the table, and its shapes with it
        else if (child.type.name === 'docCellBoxes' && !track) shapes.push(...liftedShapes(child))
        else if (child.type.name === 'docNestedTable' && child.attrs.model) {
          const model = child.attrs.model as TableModel
          if (options.nested) content.push(...nestedAsText(model, options.separator, blank))
          else content.push(state.schema.nodeFromJSON(tableModelToPmNode(model)))
        }
      })
      cells.push([...content, ...shapes])
    })
    blocks.push(...joinCells(cells, options.separator, blank))
  })
  if (blocks.length === 0) blocks.push(blank)
  const end = target.pos + target.node.nodeSize
  const tr = state.tr
  if (track) {
    // the table stays, its rows struck, in front of the text that takes its place
    const by = { author: track.author, date: revisionDateNow() }
    insertTracked(tr, end, blocks, by)
    strikeTracked(tr, target.pos + 1, end - 1, by)
    const struck = withNamesHeld(withRowRecords(tr.doc.nodeAt(target.pos)!, 'del', by))
    tr.replaceWith(target.pos, tr.mapping.map(end, -1), struck)
    tr.setSelection(TextSelection.near(tr.doc.resolve(target.pos + struck.nodeSize + 1)))
    tr.setMeta(TRACK_IGNORE, true)
  } else {
    tr.replaceWith(target.pos, end, blocks)
    tr.setSelection(TextSelection.near(tr.doc.resolve(target.pos + 1)))
  }
  editor.view.dispatch(tr.scrollIntoView())
  if (blocks.some(isCarried)) settleCarried(editor.view)
  return true
}
