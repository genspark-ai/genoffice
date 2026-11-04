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
    // PM rejects empty text nodes outright; emitting one makes setContent void
    if (node.type === 'text' && text === '') continue
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

/**
 * A NEXT template spanning every record, one page per slotsPerPage window.
 *
 * Word repeats the page until records run out — a template with 2 NEXTs fits
 * 3 records per page, so 30 records produce 10 pages. Each page is built by
 * substituteWithCursor, then pageBreakBefore stitches them together the same
 * way mergeAllRecords does for per-letter pages. Trailing pages that run past
 * the record list come out blank, matching Word's behaviour. genoffice#2097,
 * genoffice#2374.
 */
export function mergeNextPages(
  doc: PmNode,
  records: readonly MergeRecord[],
  slotsPerPage: number,
): PmNode {
  if (slotsPerPage <= 0) slotsPerPage = 1
  const content: PmNode[] = []
  const totalPages = Math.max(1, Math.ceil(records.length / slotsPerPage))
  for (let page = 0; page < totalPages; page++) {
    const startIndex = page * slotsPerPage
    const endIndex = (page + 1) * slotsPerPage - 1
    const merged = substituteWithCursor(doc, records, startIndex, endIndex)
    const pageBlocks = merged.content ?? []
    pageBlocks.forEach((block, blockIndex) => {
      const copy = page === 0 ? block : withoutUniqueAnchors(block)
      if (page === 0 || blockIndex !== 0) {
        content.push(copy)
        return
      }
      content.push({ ...copy, attrs: { ...(copy.attrs ?? {}), pageBreakBefore: true } })
    })
  }
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

/**
 * Whether the template advances the record cursor with a NEXT field, i.e. it is
 * a "one page, many records" merge rather than the "one letter per record" one.
 */
export function hasNextField(doc: PmNode): boolean {
  return countNextFields(doc) > 0
}

/**
 * The number of NEXT fields the template holds. One page holds
 * NEXT_count + 1 records (the first record, then each NEXT brings one more).
 */
export function countNextFields(doc: PmNode): number {
  let count = 0
  const walk = (node: PmNode): void => {
    const mark = node.marks?.find((m) => m.type === FIELD_MARK)
    if (mark) {
      const [keyword] = splitInstruction(String(mark.attrs?.instr ?? ''))
      if (keyword === 'NEXT') count++
    }
    for (const child of node.content ?? []) walk(child)
  }
  walk(doc)
  return count
}

/**
 * The template filled from a run of records, walked by a cursor.
 *
 * This is Word's NEXT field: the document is scanned in order, every
 * MERGEFIELD/IF takes the value of the record the cursor points at, and a NEXT
 * field advances the cursor to the next record and then disappears. The cursor
 * is held inside [startIndex, endIndex], so a template that asks for more
 * records than the window holds repeats the last one — the way Word's trailing
 * labels come out blank. genoffice#2097.
 */
export function substituteWithCursor(
  doc: PmNode,
  records: readonly MergeRecord[],
  startIndex: number,
  endIndex: number,
): PmNode {
  const cursor = { value: Math.max(startIndex, 0) }
  return { ...doc, content: walkContent(doc.content, records, cursor, endIndex) }
}

function walkContent(
  content: readonly PmNode[] | undefined,
  records: readonly MergeRecord[],
  cursor: { value: number },
  endIndex: number,
): PmNode[] {
  if (!content) return []
  const out: PmNode[] = []
  // the field the previous text node carried, and how many nodes it spans; a
  // NEXT field shares this bookkeeping so a field split over several runs only
  // advances the cursor (or fills its value) once
  let fieldInstr: string | null = null
  let fieldSeen = false
  for (const node of content) {
    const mark = node.marks?.find((m) => m.type === FIELD_MARK)
    const instr = mark ? String(mark.attrs?.instr ?? '') : ''
    const [keyword] = splitInstruction(instr)
    if (keyword === 'NEXT') {
      if (instr !== fieldInstr) {
        if (cursor.value < endIndex) cursor.value++
        fieldInstr = instr
        fieldSeen = true
      }
      // the NEXT field itself leaves nothing behind
      continue
    }
    // Word leaves trailing slots blank when the cursor runs past the window;
    // we do the same here instead of repeating the last record. genoffice#2374
    const activeRecord = cursor.value <= endIndex ? records[cursor.value] : null
    const value = mark ? fieldValue(instr, activeRecord ?? {}) : null
    if (value === null) {
      fieldInstr = null
      fieldSeen = false
      out.push(walkNode(node, records, cursor, endIndex))
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
    // PM rejects empty text nodes outright; emitting one makes setContent void
    if (node.type === 'text' && text === '') continue
    out.push({
      ...node,
      text,
      // the field is gone: what is left is the value, with the text's own marks
      marks: (node.marks ?? []).filter((m) => m.type !== FIELD_MARK),
      content: walkContent(node.content, records, cursor, endIndex),
    })
  }
  return out
}

function walkNode(
  node: PmNode,
  records: readonly MergeRecord[],
  cursor: { value: number },
  endIndex: number,
): PmNode {
  if (!node.content) return node
  return { ...node, content: walkContent(node.content, records, cursor, endIndex) }
}
