import {
  ifBranch,
  mergeFieldName,
  mergeFieldText,
  splitInstruction,
  type MergeRecord,
} from '@genoffice/docx-engine'
import type { PmNode } from '../editor/convert'

/**
 * The template half of the mail merge: the document with its merge fields
 * filled from a record.
 *
 * The editor keeps a field as a mark (instrField) on the text that shows its
 * result, so a field is found by its mark and not by a text pattern — the
 * same way the editor itself finds them. What a field is worth comes from the
 * engine (mergeFieldText), so the switches Word writes (`\b`, `\f`, `\* Upper`)
 * behave here exactly as they do in a merge.
 *
 * Nothing here knows about the UI: the panel asks for a preview and asks for
 * a merge, and this is what answers. genoffice#2097.
 */

/** the mark a field carries */
const FIELD_MARK = 'instrField'

/**
 * What a field is worth for one record, or null when the merge leaves it
 * alone.
 *
 * A MERGEFIELD takes the column's value; an IF takes the branch its condition
 * selects. Both come from the engine, so the switches Word writes behave here
 * as they do in a merge. A field that is neither (a PAGE, a DATE) is left as
 * it was: it still evaluates in Word.
 */
function fieldValue(instr: string, record: MergeRecord): string | null {
  if (mergeFieldName(instr)) return mergeFieldText(instr, record)
  const [keyword] = splitInstruction(instr)
  if (keyword === 'IF') return ifBranch(instr, record)
  return null
}

/**
 * The document with every merge field replaced by the record's value.
 *
 * A field's result may be split over several text nodes (Word does that
 * whenever the formatting changes inside it), so the nodes that carry the
 * same field are one field: the value goes into the first of them and the
 * others are emptied. Every other field (a PAGE, a DATE) is left as it was.
 */
export function substituteMergeFields(doc: PmNode, record: MergeRecord): PmNode {
  return { ...doc, content: substituteContent(doc.content, record) }
}

function substituteContent(
  content: readonly PmNode[] | undefined,
  record: MergeRecord,
): PmNode[] | undefined {
  if (!content) return undefined
  const out: PmNode[] = []
  // the field the previous text node carried, and how many nodes it spans
  let fieldInstr: string | null = null
  let fieldSeen = false
  for (const node of content) {
    const mark = node.marks?.find((m) => m.type === FIELD_MARK)
    const instr = mark ? String(mark.attrs?.instr ?? '') : ''
    const value = mark ? fieldValue(instr, record) : null
    if (value === null) {
      fieldInstr = null
      fieldSeen = false
      out.push(substituteNode(node, record))
      continue
    }
    if (instr !== fieldInstr) {
      fieldInstr = instr
      fieldSeen = false
    }
    // the field's result may span several nodes: the value goes into the
    // first of them and the others are emptied
    const text = fieldSeen ? '' : value
    fieldSeen = true
    out.push({
      ...node,
      text,
      // the field is gone: what is left is the value, with the text's own marks
      marks: (node.marks ?? []).filter((m) => m.type !== FIELD_MARK),
      content: substituteContent(node.content, record),
    })
  }
  return out
}

function substituteNode(node: PmNode, record: MergeRecord): PmNode {
  if (!node.content) return node
  return { ...node, content: substituteContent(node.content, record) }
}

/**
 * The text of a document, one line per block.
 *
 * This is the preview: the template with one record's values in it, as the
 * user reads it.
 */
export function documentText(doc: PmNode): string {
  const lines: string[] = []
  for (const block of doc.content ?? []) {
    const text = blockText(block)
    if (text !== '' || lines.length > 0) lines.push(text)
  }
  return lines.join('\n')
}

function blockText(node: PmNode): string {
  if (node.type === 'text') return node.text ?? ''
  if (!node.content) return ''
  return node.content.map(blockText).join('')
}

/**
 * Every merge field the template holds, in document order, without repeats.
 *
 * The panel offers these as the fields it can insert, and the preview tells
 * the user which of them the current record fills.
 */
export function mergeFieldNames(doc: PmNode): string[] {
  const names: string[] = []
  const walk = (node: PmNode): void => {
    for (const mark of node.marks ?? []) {
      const name = mergeFieldName(String(mark.attrs?.instr ?? ''))
      if (name && !names.includes(name)) names.push(name)
    }
    for (const child of node.content ?? []) walk(child)
  }
  walk(doc)
  return names
}

/**
 * Every record's letter in one document.
 *
 * Word's "Edit Individual Documents" puts the letters one after another, each
 * on its own page; the same shape is built here: the template's blocks
 * repeated per record, with the page break on the first block of every
 * letter but the first.
 */
export function mergeAllRecords(doc: PmNode, records: readonly MergeRecord[]): PmNode {
  const blocks = doc.content ?? []
  const content: PmNode[] = []
  records.forEach((record, index) => {
    const merged = substituteContent(blocks, record) ?? []
    merged.forEach((block, blockIndex) => {
      // a bookmark, a comment anchor or a note reference exists once in the
      // package, so only the first letter keeps it
      const copy = index === 0 ? block : withoutUniqueAnchors(block)
      if (index === 0 || blockIndex !== 0) {
        content.push(copy)
        return
      }
      content.push({ ...copy, attrs: { ...(copy.attrs ?? {}), pageBreakBefore: true } })
    })
  })
  return { ...doc, content }
}

function withoutUniqueAnchors(node: PmNode): PmNode {
  const out: PmNode = { ...node }
  if (out.attrs && 'bookmarks' in out.attrs && out.attrs.bookmarks) {
    out.attrs = { ...out.attrs, bookmarks: null }
  }
  if (out.marks?.some((m) => m.type === 'comment')) {
    out.marks = out.marks.filter((m) => m.type !== 'comment')
  }
  if (node.content) {
    out.content = node.content
      .filter((child) => child.type !== 'docNoteRef')
      .map(withoutUniqueAnchors)
  }
  return out
}
