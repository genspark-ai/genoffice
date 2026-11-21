/** Word's Table > Convert: paragraphs into a table's cells, and a table back into paragraphs */
import type { Editor } from '@tiptap/core'
import type { TableAutoFitMode, TableModel } from '@genoffice/docx-engine'
import { Fragment, type Node as PmNode } from '@tiptap/pm/model'
import { TextSelection, type EditorState } from '@tiptap/pm/state'
import { tableModelToPmNode } from './convert'
import { wordTableCell, wordTableModel } from './new-table'
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

function cellFragments(blocks: readonly Block[], separator: ConvertSeparator): Fragment[][] {
  return blocks.map(({ node }) =>
    byParagraph(separator) ? [node.content] : fieldsOf(node, separator.text),
  )
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
  const perBlock = cellFragments(blocks, options.separator)
  const rows: Fragment[][] = []
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
      const content = rows[rowIndex]?.[cellIndex]
      const paragraph = cell.firstChild
      cells.push(content && paragraph ? cell.copy(Fragment.from(paragraph.copy(content))) : cell)
    })
    filled.push(row.copy(Fragment.from(cells)))
  })
  const table = empty.copy(Fragment.from(filled))
  const from = blocks[0].pos
  const last = blocks[blocks.length - 1]
  const to = last.pos + last.node.nodeSize
  const tr = state.tr.replaceWith(from, to, table)
  tr.setSelection(TextSelection.near(tr.doc.resolve(from + 1)))
  editor.view.dispatch(tr.scrollIntoView())
  if (options.autoFit !== 'window') {
    // rides the same undo step as the conversion
    setTableAutoFit(options.autoFit, contentWidthPx)(editor.state, (fit) =>
      editor.view.dispatch(fit.setMeta('appendedTransaction', tr)),
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

/** one row's cells in a line: the last paragraph of a cell runs on into the first of the next */
function joinCells(cells: PmNode[][], separator: ConvertSeparator, paragraph: PmNode): PmNode[] {
  if (separator.kind === 'paragraph') return cells.flat()
  const text = separator.text ? paragraph.type.schema.text(separator.text) : null
  const line: PmNode[] = []
  cells.forEach((blocks, index) => {
    const [first, ...rest] = blocks.length > 0 ? blocks : [paragraph]
    const prev = line[line.length - 1]
    if (index === 0 || !prev?.isTextblock || !first.isTextblock) {
      line.push(first, ...rest)
      return
    }
    const joined = text ? prev.content.addToEnd(text) : prev.content
    line[line.length - 1] = prev.copy(joined.append(first.content))
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
  target.node.forEach((row) => {
    const cells: PmNode[][] = []
    row.forEach((cell) => {
      if (!isCell(cell) || cell.attrs.gridGap || cell.attrs.mergeGap) return
      const content: PmNode[] = []
      cell.forEach((child) => {
        if (child.isTextblock) content.push(child)
        else if (child.type.name === 'docNestedTable' && child.attrs.model) {
          const model = child.attrs.model as TableModel
          if (options.nested) content.push(...nestedAsText(model, options.separator, blank))
          else content.push(state.schema.nodeFromJSON(tableModelToPmNode(model)))
        }
      })
      cells.push(content)
    })
    blocks.push(...joinCells(cells, options.separator, blank))
  })
  if (blocks.length === 0) blocks.push(blank)
  const tr = state.tr.replaceWith(target.pos, target.pos + target.node.nodeSize, blocks)
  tr.setSelection(TextSelection.near(tr.doc.resolve(target.pos + 1)))
  editor.view.dispatch(tr.scrollIntoView())
  return true
}
