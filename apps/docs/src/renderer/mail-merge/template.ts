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
 * NEXTIF / SKIPIF carry a condition this merge does not evaluate. Their result
 * text is Word's own instruction placeholder, not a value, so a merged letter
 * drops the node rather than printing the instruction. (Full condition
 * evaluation would also decide when they advance/skip a record.)
 */
function isUnevaluatedRule(instr: string): boolean {
  const [keyword] = splitInstruction(instr)
  return keyword === 'NEXTIF' || keyword === 'SKIPIF'
}

/**
 * How many pages a merge of recordCount records takes, one template copy
 * holding slotsPerPage records. Zero records still write nothing.
 */
export function mergePageCount(recordCount: number, slotsPerPage: number): number {
  if (recordCount <= 0) return 0
  const slots = slotsPerPage > 0 ? slotsPerPage : 1
  return Math.ceil(recordCount / slots)
}

/**
 * The page a previewed record sits on: the window "current record" stands for.
 *
 * Preview fills one record per copy, but the merge writes the whole page, so
 * the window is the page group — the same arithmetic mergeNextPages and the
 * output's per-page files use (the index decides the page, slot 1 is the
 * previewed record).
 */
export function recordPage(
  recordCount: number,
  index: number,
  slotsPerPage: number,
): { start: number; end: number; count: number } {
  if (recordCount <= 0) return { start: 0, end: -1, count: 0 }
  const slots = slotsPerPage > 0 ? slotsPerPage : 1
  const lastRecord = recordCount - 1
  const start = Math.min(Math.floor(Math.max(index, 0) / slots) * slots, lastRecord)
  const end = Math.min(start + slots - 1, lastRecord)
  return { start, end, count: end - start + 1 }
}

/**
 * The template with every merge field replaced by the record's value.
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
    // an unevaluated NEXTIF/SKIPIF leaves no instruction text behind
    if (mark && isUnevaluatedRule(instr)) {
      fieldInstr = null
      fieldSeen = false
      continue
    }
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
  // nothing to merge: the template as-is, not one blank trailing page
  if (records.length === 0) return doc
  if (slotsPerPage <= 0) slotsPerPage = 1
  const content: PmNode[] = []
  const totalPages = Math.max(1, mergePageCount(records.length, slotsPerPage))
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
    // one NEXT field split over several runs advances the cursor once, so it
    // must be counted once — otherwise the page window (1 + count) would be
    // wider than the template's real slots and records would be skipped. A run
    // of adjacent NEXT nodes is one field even when Word split it into runs
    // with slightly different instruction text.
    let lastWasNext = false
    for (const child of node.content ?? []) {
      const mark = child.marks?.find((m) => m.type === FIELD_MARK)
      const instr = mark ? String(mark.attrs?.instr ?? '') : ''
      const [keyword] = splitInstruction(instr)
      if (keyword === 'NEXT') {
        if (!lastWasNext) count++
        lastWasNext = true
        continue
      }
      lastWasNext = false
      walk(child)
    }
  }
  walk(doc)
  return count
}

/**
 * Whether a node renders something: non-whitespace text, or a rendered child.
 *
 * An empty paragraph, a paragraph mark, a page/section break and an empty
 * table row carry no content, so they cannot make a page hold another record.
 */
function hasVisibleContent(node: PmNode): boolean {
  if (node.type === 'text') return (node.text ?? '').trim() !== ''
  if (node.type === 'docPageBreak' || node.type === 'docSectionBreak') return false
  return (node.content ?? []).some(hasVisibleContent)
}

/**
 * How many records one copy of the template consumes.
 *
 * This is not NEXT_count + 1: a NEXT at the end of the page only hands the
 * cursor to the next page's first block, it renders nothing, so it adds no
 * slot. Counting NEXT fields instead would widen the page window past the
 * blocks that exist and skip a record between pages.
 */
export function countRecordSlots(doc: PmNode): number {
  let cursor = 0
  let highest = -1
  const walk = (content: readonly PmNode[] | undefined): void => {
    if (!content) return
    // the same field bookkeeping walkContent uses, so a field split over
    // several runs is counted once
    let fieldInstr: string | null = null
    let fieldSeen = false
    let lastWasNext = false
    for (const node of content) {
      const mark = node.marks?.find((m) => m.type === FIELD_MARK)
      const instr = mark ? String(mark.attrs?.instr ?? '') : ''
      const [keyword] = splitInstruction(instr)
      if (keyword === 'NEXT') {
        if (!lastWasNext) {
          cursor++
        }
        lastWasNext = true
        continue
      }
      lastWasNext = false
      const value = mark ? fieldValue(instr, {}) : null
      if (value === null) {
        fieldInstr = null
        fieldSeen = false
        walk(node.content)
        continue
      }
      // a blank field or a whitespace-only run renders nothing, so the block
      // it sits in is not a record block (Word leaves NEXT/IF placeholders on
      // otherwise empty lines)
      if (!hasVisibleContent(node)) {
        fieldInstr = null
        fieldSeen = false
        continue
      }
      if (instr !== fieldInstr) {
        fieldInstr = instr
        fieldSeen = false
      }
      // a field that renders takes the record the cursor points at; the other
      // nodes of the same field reuse it
      if (!fieldSeen) highest = Math.max(highest, cursor)
      fieldSeen = true
      walk(node.content)
    }
  }
  walk(doc.content)
  // no field at all: one page still holds one letter
  return Math.max(1, highest + 1)
}

/** one template block's part in the slot count, for the diagnostics below */
export interface SlotBlockTrace {
  index: number
  type: string
  visible: boolean
  /** flattened, whitespace-collapsed text of the block (first 60 chars) */
  text: string
  nextFields: number
  fields: string[]
  /** highest slot occupied before this block was walked (0 = none yet) */
  slotBefore: number
  /** highest slot occupied after this block was walked */
  slotAfter: number
  contributed: boolean
}

/** the reason a page holds N records, block by block */
export interface SlotTrace {
  slots: number
  blocks: SlotBlockTrace[]
  nextFields: Array<{ block: number; text: string; instr: string }>
}

/**
 * DEV diagnostics only: explain which template block takes each record slot.
 *
 * Mirrors countRecordSlots' walk but reports at top-level block granularity (a
 * table's rows collapse into the table's block), so the finish dialog can say
 * why a page holds N records. The merge itself never calls this.
 */
export function describeRecordSlots(doc: PmNode): SlotTrace {
  const blocks: SlotBlockTrace[] = []
  const nextFields: SlotTrace['nextFields'] = []
  let cursor = 0
  let highest = -1
  const topBlocks = doc.content ?? []
  for (let index = 0; index < topBlocks.length; index++) {
    const block = topBlocks[index]
    const slotBefore = highest + 1
    const occurrences: Array<{ instr: string; text: string }> = []
    let lastWasNext = false
    const scan = (node: PmNode): void => {
      const mark = node.marks?.find((m) => m.type === FIELD_MARK)
      const instr = mark ? String(mark.attrs?.instr ?? '') : ''
      const [keyword] = splitInstruction(instr)
      if (keyword === 'NEXT') {
        if (!lastWasNext) cursor++
        lastWasNext = true
        nextFields.push({ block: index, text: node.text ?? '', instr })
        return
      }
      lastWasNext = false
      if (mark && fieldValue(instr, {}) !== null) occurrences.push({ instr, text: node.text ?? '' })
      for (const child of node.content ?? []) scan(child)
    }
    for (const child of block.content ?? []) scan(child)
    const visible = hasVisibleContent(block)
    let prevInstr: string | null = null
    let seen = false
    for (const occurrence of occurrences) {
      // countRecordSlots resets its field bookkeeping on an invisible block
      if (!visible) {
        prevInstr = null
        seen = false
        continue
      }
      if (occurrence.instr !== prevInstr) {
        prevInstr = occurrence.instr
        seen = false
      }
      if (!seen) highest = Math.max(highest, cursor)
      seen = true
    }
    blocks.push({
      index,
      type: block.type,
      visible,
      text: blockText(block).replace(/\s+/g, ' ').trim().slice(0, 60),
      nextFields: nextFields.filter((entry) => entry.block === index).length,
      fields: occurrences.map((entry) => entry.instr),
      slotBefore,
      slotAfter: highest + 1,
      contributed: highest + 1 > slotBefore,
    })
  }
  return { slots: countRecordSlots(doc), blocks, nextFields }
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
    // NEXTIF/SKIPIF are not evaluated here; at least their instruction text
    // must not survive into the merged letter (it is not a record value)
    if (mark && (keyword === 'NEXTIF' || keyword === 'SKIPIF')) {
      fieldInstr = null
      fieldSeen = false
      continue
    }
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
