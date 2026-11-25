import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'

/** what Word's Mark Index Entry dialog takes: the entry and where the index sends the reader */
export interface IndexEntry {
  main: string
  sub: string
  /** the page the mark is on, a cross-reference in its place, or the pages of a bookmark */
  target: 'page' | 'cross' | 'range'
  cross: string
  bookmark: string
  bold: boolean
  italic: boolean
}

/** what an XE field's switches say */
export interface IndexEntryOptions {
  /** `\t`: the text the index shows in place of a page */
  cross?: string
  /** `\r`: the bookmark whose pages the index lists */
  bookmark?: string
  bold: boolean
  italic: boolean
}

// a quote would end the field's text early
const literal = (text: string) => text.replace(/"/g, '').trim()

/** the entry text of an XE field: "Main:Sub" */
export const indexEntryTerm = (entry: Pick<IndexEntry, 'main' | 'sub'>): string =>
  [literal(entry.main), literal(entry.sub)].filter(Boolean).join(':')

/** the switches of an XE field for the entry; null when it has none */
export function indexEntrySwitches(entry: IndexEntry): string | null {
  const parts: string[] = []
  if (entry.target === 'cross' && literal(entry.cross)) parts.push(`\\t "${literal(entry.cross)}"`)
  if (entry.target === 'range' && entry.bookmark) parts.push(`\\r ${entry.bookmark}`)
  if (entry.bold) parts.push('\\b')
  if (entry.italic) parts.push('\\i')
  return parts.length > 0 ? parts.join(' ') : null
}

export function indexEntryOptions(switches: string | null | undefined): IndexEntryOptions {
  const text = switches ?? ''
  const cross = /\\t\s+"([^"]*)"/i.exec(text)?.[1]
  const bookmark = /\\r\s+"?([^\s"\\]+)"?/i.exec(text)?.[1]
  // the switches after a quoted text are the bare ones
  const bare = text.replace(/"[^"]*"/g, '')
  return {
    ...(cross ? { cross } : {}),
    ...(bookmark ? { bookmark } : {}),
    bold: /\\b(?![a-z])/i.test(bare),
    italic: /\\i(?![a-z])/i.test(bare),
  }
}

/** an inline node counts as one character, as it does in a position */
const OBJECT = '\ufffc'

/**
 * Where a mark for the occurrences of the text goes, capitals and all: after
 * the first one in each paragraph (Word's Mark All for an index entry), or
 * after every one (its Mark All for a citation).
 */
export function occurrenceEnds(editor: Editor, text: string, every = false): number[] {
  const out: number[] = []
  if (!text) return out
  editor.state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return
    let flat = ''
    node.forEach((child) => {
      flat += child.isText ? (child.text ?? '') : OBJECT.repeat(child.nodeSize)
    })
    for (let at = flat.indexOf(text); at !== -1; at = flat.indexOf(text, at + text.length)) {
      out.push(pos + 1 + at + text.length)
      if (!every) break
    }
    return false
  })
  return out
}

/** puts an index entry or citation marker with those attributes at each position, as one step */
export function insertFieldMarks(
  editor: Editor,
  marks: Array<{ at: number; attrs: Record<string, unknown> }>,
): void {
  const type = editor.state.schema.nodes.docXeMark
  if (!type || marks.length === 0) return
  const { tr, doc } = editor.state
  for (const mark of [...marks].sort((a, b) => b.at - a.at))
    tr.insert(Math.min(mark.at, doc.content.size), type.create(mark.attrs))
  // the caret goes behind what was selected: the next thing typed or inserted must not replace it
  tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(editor.state.selection.to)))
  editor.view.dispatch(tr)
  editor.commands.focus()
}

/** puts an index entry (XE field) at each position */
export function markIndexEntries(
  editor: Editor,
  positions: number[],
  term: string,
  switches: string | null,
): void {
  if (term)
    insertFieldMarks(
      editor,
      positions.map((at) => ({ at, attrs: { term, switches } })),
    )
}

/** what Word's Mark Table of Contents Entry dialog takes */
export interface TocEntry {
  text: string
  /** the table the entry belongs to: a letter, C being the table of contents itself */
  id: string
  level: number
}

/** the text and switches of the entry's TC field; an entry of the default table names no table */
export function tocEntryInstruction(entry: TocEntry): string {
  const id = /^[A-Za-z]$/.test(entry.id) ? entry.id.toUpperCase() : 'C'
  const level = Math.min(9, Math.max(1, Math.round(entry.level) || 1))
  return `"${literal(entry.text)}"${id === 'C' ? '' : ` \\f ${id}`} \\l ${level}`
}

/** puts a table of contents entry (TC field) at the position */
export function markTocEntry(editor: Editor, at: number, entry: TocEntry): void {
  if (literal(entry.text))
    insertFieldMarks(editor, [{ at, attrs: { tc: tocEntryInstruction(entry) } }])
}
